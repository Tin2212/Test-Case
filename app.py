import os
import json
import pandas as pd
import io
from types import SimpleNamespace
from urllib.parse import quote 
from flask import (Flask, render_template, request, redirect, url_for, 
                   flash, Response, send_from_directory)
from sqlalchemy import func, or_
from werkzeug.utils import secure_filename
import uuid

# 從我們的新檔案中匯入所需物件
from extensions import db, migrate 
from models import TestCase, Tag, Attachment
from services import process_excel_file
from utils import categorize_case, process_tags, load_category_rules

# --- 1. 初始化與設定 ---
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
ATTACHMENT_FOLDER = os.path.join(UPLOAD_FOLDER, 'attachments') 
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['SECRET_KEY'] = 'a_very_secure_and_random_secret_key_for_production'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///testcases.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['ATTACHMENT_FOLDER'] = ATTACHMENT_FOLDER 

db.init_app(app)
migrate.init_app(app, db)

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(ATTACHMENT_FOLDER):
    os.makedirs(ATTACHMENT_FOLDER)

# --- 路由 (Web 頁面邏輯) ---

# ... (context_processor, dashboard, search, index, add_case, edit_case, delete_case, status_result 相關路由保持不變) ...
@app.context_processor
def inject_status_options():
    status_options = ['未執行', '進行中', '通過', '失敗']
    return dict(status_options=status_options)

@app.route('/dashboard')
def dashboard():
    status_distribution = db.session.query(
        TestCase.status,
        func.count(TestCase.status)
    ).group_by(TestCase.status).all()
    
    pie_chart_data = {
        'labels': [status[0] for status in status_distribution],
        'data': [status[1] for status in status_distribution]
    }

    main_categories = db.session.query(
        TestCase.main_category
    ).group_by(TestCase.main_category).order_by(TestCase.main_category).all()
    
    progress_data = []
    for category_tuple in main_categories:
        category_name = category_tuple[0]
        if not category_name:
            continue

        total = TestCase.query.filter_by(main_category=category_name).count()
        completed = TestCase.query.filter(
            TestCase.main_category == category_name,
            TestCase.status.in_(['通過', '失敗'])
        ).count()
        passed = TestCase.query.filter_by(main_category=category_name, status='通過').count()

        completion_percentage = (completed / total * 100) if total > 0 else 0
        pass_percentage = (passed / total * 100) if total > 0 else 0

        progress_data.append({
            'category': category_name.replace('功能', ''),
            'total': total,
            'completed': completed,
            'passed': passed,
            'completion_percentage': round(completion_percentage, 1),
            'pass_percentage': round(pass_percentage, 1)
        })

    summary_data = {
        'total_cases': sum(item['total'] for item in progress_data),
        'completed_cases': sum(item['completed'] for item in progress_data),
        'passed_cases': sum(item['passed'] for item in progress_data)
    }

    return render_template(
        'dashboard.html',
        pie_chart_data=json.dumps(pie_chart_data),
        progress_data=progress_data,
        summary_data=summary_data,
        hide_sidebar=True
    )

@app.route('/search')
def search():
    query_term = request.args.get('q', '').strip()
    
    if not query_term:
        return redirect(url_for('index'))

    search_filter = or_(
        TestCase.case_id.ilike(f'%{query_term}%'),
        TestCase.test_item.ilike(f'%{query_term}%'),
        TestCase.test_purpose.ilike(f'%{query_term}%'),
        TestCase.test_steps.ilike(f'%{query_term}%'),
        TestCase.notes.ilike(f'%{query_term}%')
    )

    found_cases = TestCase.query.filter(search_filter).order_by(TestCase.case_id).all()

    mock_pagination = SimpleNamespace(
        total=len(found_cases),
        pages=0,
        iter_pages=lambda: []
    )

    return render_template(
        'search_results.html', 
        cases=found_cases, 
        query_term=query_term,
        hide_sidebar=True,
        tree_data={},
        pagination=mock_pagination
    )

@app.route('/')
def index():
    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 50, type=int)
    if per_page not in [10, 20, 30, 40, 50]:
        per_page = 50
    
    selected_statuses = request.args.getlist('status')
    selected_tags = request.args.getlist('tag')

    all_cases_for_tree = db.session.query(TestCase.product_type, TestCase.main_category, TestCase.sub_category).distinct().all()
    tree_data = {}
    for prod, main_cat, sub_cat in all_cases_for_tree:
        if not prod: continue
        if prod not in tree_data: tree_data[prod] = {}
        if main_cat and main_cat not in tree_data[prod]: tree_data[prod][main_cat] = []
        if main_cat and sub_cat and sub_cat not in tree_data[prod][main_cat]:
            tree_data[prod][main_cat].append(sub_cat)
    
    selected_product = request.args.get('product')
    selected_main_category = request.args.get('main_category')
    selected_sub_category = request.args.get('sub_category')

    global_precondition = None
    if selected_product:
        CATEGORY_RULES = load_category_rules()
        global_preconditions = CATEGORY_RULES.get('global_preconditions', {})
        global_precondition = global_preconditions.get(selected_product)

    query = TestCase.query
    if selected_product:
        query = query.filter_by(product_type=selected_product)
    if selected_main_category:
        query = query.filter_by(main_category=selected_main_category)
    if selected_sub_category:
        query = query.filter_by(sub_category=selected_sub_category)

    if selected_statuses:
        query = query.filter(TestCase.status.in_(selected_statuses))
    if selected_tags:
        for tag_name in selected_tags:
            query = query.filter(TestCase.tags.any(name=tag_name))

    pagination = query.order_by(TestCase.case_id).paginate(page=page, per_page=per_page, error_out=False)
    cases_to_display = pagination.items
    
    all_tags = Tag.query.order_by(Tag.name).all()

    return render_template('cases.html', 
                           cases=cases_to_display, 
                           tree_data=tree_data,
                           pagination=pagination,
                           selected_product=selected_product, 
                           selected_main_category=selected_main_category,
                           selected_sub_category=selected_sub_category,
                           global_precondition=global_precondition,
                           per_page=per_page,
                           all_tags=all_tags,
                           selected_tags=selected_tags,
                           selected_statuses=selected_statuses)

@app.route('/add', methods=['GET', 'POST'])
def add_case():
    if request.method == 'POST':
        case_data = request.form.to_dict()
        product_type = case_data.get('product_type', '未分類產品')
        main_cat, sub_cat = categorize_case(case_data, product_type)
        
        new_case = TestCase(
            product_type=case_data.get('product_type', '未分類產品'),
            category=case_data.get('category', ''),
            main_category=main_cat, sub_category=sub_cat,
            case_id=case_data.get('case_id'), test_item=case_data.get('test_item'),
            test_purpose=case_data.get('test_purpose'), preconditions=case_data.get('preconditions'),
            test_steps=case_data.get('test_steps'), expected_result=case_data.get('expected_result'),
            actual_result=case_data.get('actual_result'), status=case_data.get('status', '未執行'),
            notes=case_data.get('notes'), reference=case_data.get('reference')
        )
        tags_string = case_data.get('tags', '')
        new_case.tags = process_tags(tags_string)

        db.session.add(new_case)
        db.session.commit()
        flash('測試案例已成功新增！', 'success')
        return redirect(url_for('index'))
    return render_template('case_form.html', title="新增測試案例", case=None)

@app.route('/edit/<int:id>', methods=['GET', 'POST'])
def edit_case(id):
    case_to_edit = TestCase.query.get_or_404(id)
    if request.method == 'POST':
        case_data = request.form.to_dict()
        product_type = case_data.get('product_type')
        main_cat, sub_cat = categorize_case(case_data, product_type)
        
        case_to_edit.product_type = case_data.get('product_type')
        case_to_edit.category = case_data.get('category')
        case_to_edit.main_category = main_cat
        case_to_edit.sub_category = sub_cat
        case_to_edit.case_id = case_data.get('case_id')
        case_to_edit.test_item = case_data.get('test_item')
        case_to_edit.test_purpose = case_data.get('test_purpose')
        case_to_edit.preconditions = case_data.get('preconditions')
        case_to_edit.test_steps = case_data.get('test_steps')
        case_to_edit.expected_result = case_data.get('expected_result')
        case_to_edit.actual_result = case_data.get('actual_result')
        case_to_edit.status = case_data.get('status')
        case_to_edit.notes = case_data.get('notes')
        case_to_edit.reference = case_data.get('reference')
        
        tags_string = case_data.get('tags', '')
        case_to_edit.tags = process_tags(tags_string)

        db.session.commit()
        flash('測試案例已成功更新！', 'success')
        return redirect(url_for('index'))
    return render_template('case_form.html', title="編輯測試案例", case=case_to_edit)

@app.route('/delete/<int:id>', methods=['POST'])
def delete_case(id):
    case_to_delete = TestCase.query.get_or_404(id)
    db.session.delete(case_to_delete)
    db.session.commit()
    return '', 200

@app.route('/edit-status-result/<int:id>', methods=['GET', 'POST'])
def edit_status_result(id):
    case = TestCase.query.get_or_404(id)
    if request.method == 'POST':
        case.status = request.form.get('status')
        case.actual_result = request.form.get('actual_result', '')
        db.session.commit()
        return render_template('partials/_status_result_display.html', case=case)
    return render_template('partials/_status_result_edit.html', case=case)

@app.route('/display-status-result/<int:id>')
def display_status_result(id):
    case = TestCase.query.get_or_404(id)
    return render_template('partials/_status_result_display.html', case=case)

@app.route('/delete-tag/<int:id>', methods=['POST'])
def delete_tag(id):
    case = TestCase.query.get_or_404(id)
    tag_name_to_delete = request.form.get('tag')
    if tag_name_to_delete:
        tag_to_delete = Tag.query.filter_by(name=tag_name_to_delete).first()
        if tag_to_delete and tag_to_delete in case.tags:
            case.tags.remove(tag_to_delete)
            db.session.commit()
    return render_template('partials/_tags_display.html', case=case)

@app.route('/bulk-add-tag', methods=['POST'])
def bulk_add_tag():
    case_ids = request.form.getlist('case_ids')
    new_tag_name = request.form.get('new_tag', '').strip().lower()

    redirect_params = {
        'product': request.form.get('product'),
        'main_category': request.form.get('main_category'),
        'sub_category': request.form.get('sub_category'),
        'page': request.form.get('page', 1, type=int),
        'per_page': request.form.get('per_page', 50, type=int)
    }
    cleaned_redirect_params = {k: v for k, v in redirect_params.items() if v}

    if not case_ids or not new_tag_name:
        flash('未選擇任何案例或未輸入標籤。', 'warning')
        return redirect(url_for('index', **cleaned_redirect_params))

    tag_to_add = Tag.query.filter_by(name=new_tag_name).first()
    if not tag_to_add:
        tag_to_add = Tag(name=new_tag_name)
        db.session.add(tag_to_add)

    cases_to_update = TestCase.query.filter(TestCase.id.in_(case_ids)).all()
    for case in cases_to_update:
        if tag_to_add not in case.tags:
            case.tags.append(tag_to_add)
            
    db.session.commit()
    flash(f'已為 {len(case_ids)} 個案例成功新增標籤 "{new_tag_name}"！', 'success')
    return redirect(url_for('index', **cleaned_redirect_params))

@app.route('/bulk-delete', methods=['POST'])
def bulk_delete():
    case_ids = request.form.getlist('case_ids')
    
    redirect_params = {
        'product': request.form.get('product'),
        'main_category': request.form.get('main_category'),
        'sub_category': request.form.get('sub_category'),
        'page': request.form.get('page', 1, type=int),
        'per_page': request.form.get('per_page', 50, type=int)
    }
    cleaned_redirect_params = {k: v for k, v in redirect_params.items() if v}

    if not case_ids:
        flash('未選擇任何案例。', 'warning')
        return redirect(url_for('index', **cleaned_redirect_params))

    cases_to_delete = TestCase.query.filter(TestCase.id.in_(case_ids)).all()
    
    for case in cases_to_delete:
        db.session.delete(case)
        
    db.session.commit()
    
    flash(f'已成功刪除 {len(cases_to_delete)} 個案例！', 'success')
    return redirect(url_for('index', **cleaned_redirect_params))

@app.route('/edit-notes/<int:id>', methods=['GET', 'POST'])
def edit_notes(id):
    case = TestCase.query.get_or_404(id)
    if request.method == 'POST':
        case.notes = request.form.get('notes', '')
        
        if 'attachment' in request.files:
            file = request.files['attachment']
            if file and file.filename != '':
                original_filename = secure_filename(file.filename)
                unique_filename = f"{uuid.uuid4().hex}_{original_filename}"
                file_path = os.path.join(app.config['ATTACHMENT_FOLDER'], unique_filename)
                file.save(file_path)
                
                new_attachment = Attachment(
                    filename=original_filename,
                    filepath=unique_filename,
                    test_case_id=case.id
                )
                db.session.add(new_attachment)

        db.session.commit()
        
        # ★★★ 核心修正點 2: 回傳 display 樣板，並加上觸發器 ★★★
        response = Response(render_template('partials/_notes_display.html', case=case))
        response.headers['HX-Trigger'] = f'refreshDetails-{case.id}'
        return response

    return render_template('partials/_notes_edit.html', case=case)

@app.route('/display-notes/<int:id>')
def display_notes(id):
    case = TestCase.query.get_or_404(id)
    return render_template('partials/_notes_display.html', case=case)

@app.route('/attachments/view/<filename>')
def serve_attachment(filename):
    return send_from_directory(app.config['ATTACHMENT_FOLDER'], filename)

@app.route('/attachments/download/<filename>')
def download_attachment(filename):
    return send_from_directory(app.config['ATTACHMENT_FOLDER'], filename, as_attachment=True)

@app.route('/attachments/delete/<int:attachment_id>', methods=['POST'])
def delete_attachment(attachment_id):
    attachment = Attachment.query.get_or_404(attachment_id)
    case_id = attachment.test_case_id
    
    try:
        os.remove(os.path.join(app.config['ATTACHMENT_FOLDER'], attachment.filepath))
    except OSError as e:
        print(f"Error deleting file {attachment.filepath}: {e}")

    db.session.delete(attachment)
    db.session.commit()

    case = TestCase.query.get_or_404(case_id)
    
    # ★★★ 核心修正點 3: 回傳 edit 樣板，並加上觸發器 ★★★
    response = Response(render_template('partials/_notes_edit.html', case=case))
    response.headers['HX-Trigger'] = f'refreshDetails-{case.id}'
    return response

# ★★★ 核心修正點 4: 新增專門用來刷新詳細內容的路由 ★★★
@app.route('/case-details/<int:id>')
def get_case_details(id):
    case = TestCase.query.get_or_404(id)
    # 記得要傳遞 render_manual_list 函式給 partial
    return render_template('partials/_case_details_content.html', case=case, render_manual_list=render_manual_list)

# 輔助函式，讓樣板可以呼叫
@app.context_processor
def utility_processor():
    def render_manual_list_in_template(text_block):
        lines = [line.strip().lstrip('0123456789. ') for line in (text_block or "").split('\n') if line.strip()]
        html = '<div class="manual-list">'
        for i, line in enumerate(lines):
            html += f'<div class="manual-list-item"><span class="manual-list-number">{i+1}.</span><span class="manual-list-text">{line}</span></div>'
        html += '</div>'
        return html
    return dict(render_manual_list=render_manual_list_in_template)


@app.route('/upload', methods=['GET', 'POST'])
def upload_page():
    if request.method == 'POST':
        uploaded_files = request.files.getlist('files')
        if not uploaded_files or uploaded_files[0].filename == '':
            flash('未選擇任何檔案', 'warning')
            return redirect(request.url)

        total_imported_count = 0
        has_error = False

        for file in uploaded_files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                try:
                    count = process_excel_file(file.stream, filename)
                    total_imported_count += count
                except Exception as e:
                    has_error = True
                    flash(f'處理檔案 "{filename}" 時發生錯誤：{e}', 'danger')
                    db.session.rollback()
                    break 

        if not has_error and total_imported_count > 0:
             flash(f'所有檔案處理完畢！共成功匯入 {total_imported_count} 筆新案例！', 'success')
        elif not has_error and total_imported_count == 0:
            flash('所有檔案處理完畢，但沒有匯入任何新案例 (可能 Case ID 皆已存在)。', 'info')

        return redirect(url_for('index'))

    return render_template('upload.html')

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/export')
def export_cases():
    product = request.args.get('product')
    main_category = request.args.get('main_category')
    sub_category = request.args.get('sub_category')
    
    selected_statuses = request.args.getlist('status')
    selected_tags = request.args.getlist('tag')

    query = TestCase.query
    if product:
        query = query.filter_by(product_type=product)
    if main_category:
        query = query.filter_by(main_category=main_category)
    if sub_category:
        query = query.filter_by(sub_category=sub_category)
    if selected_statuses:
        query = query.filter(TestCase.status.in_(selected_statuses))
    if selected_tags:
        for tag_name in selected_tags:
            query = query.filter(TestCase.tags.any(name=tag_name))

    cases_to_export = query.order_by(TestCase.case_id).all()

    if not cases_to_export:
        flash('沒有符合目前篩選條件的資料可供匯出。', 'warning')
        return redirect(request.referrer or url_for('index'))

    data_for_df = [{
        'Case ID': case.case_id,
        '產品類型': case.product_type,
        '主分類': case.main_category.replace('功能', '') if case.main_category else '',
        '子分類': case.sub_category,
        '測試項目': case.test_item,
        '測試目的': case.test_purpose,
        '前置條件': case.preconditions,
        '測試步驟': case.test_steps,
        '預期結果': case.expected_result,
        '實際結果': case.actual_result,
        '狀態': case.status,
        '標籤': ", ".join(tag.name for tag in case.tags),
        '備註': case.notes
    } for case in cases_to_export]
    df = pd.DataFrame(data_for_df)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='TestCases')
        worksheet = writer.sheets['TestCases']
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)
    output.seek(0)

    filename = "test_cases_export.xlsx"
    if sub_category:
        filename = f"{sub_category}.xlsx"
    elif main_category:
        filename = f"{main_category.replace('功能', '')}.xlsx"
    elif product:
        filename = f"{product}.xlsx"
    
    return Response(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"
        }
    )

# --- 5. 啟動程式 ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)