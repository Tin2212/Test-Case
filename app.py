import os
import json
import pandas as pd
from types import SimpleNamespace
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func, or_
from flask_migrate import Migrate
from werkzeug.utils import secure_filename

# --- 1. 初始化與設定 ---
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['SECRET_KEY'] = 'a_very_secure_and_random_secret_key_for_production'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///testcases.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

db = SQLAlchemy(app)
migrate = Migrate(app, db)

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- 2. 資料庫模型定義 (標籤系統升級) ---

# 步驟 1: 建立 TestCase 和 Tag 之間的多對多關聯表
test_case_tags = db.Table('test_case_tags',
    db.Column('test_case_id', db.Integer, db.ForeignKey('test_case.id'), primary_key=True),
    db.Column('tag_id', db.Integer, db.ForeignKey('tag.id'), primary_key=True)
)

# 步驟 2: 建立新的 Tag 模型
class Tag(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)

    def __repr__(self):
        return f'<Tag {self.name}>'

class TestCase(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_type = db.Column(db.String(50), nullable=False, default='未分類產品')
    category = db.Column(db.String(100), nullable=False)
    main_category = db.Column(db.String(50), nullable=True)
    sub_category = db.Column(db.String(50), nullable=True)
    case_id = db.Column(db.String(50), unique=True, nullable=False)
    test_item = db.Column(db.String(200), nullable=False)
    test_purpose = db.Column(db.Text, nullable=True)
    preconditions = db.Column(db.Text, nullable=True)
    test_steps = db.Column(db.Text, nullable=True)
    expected_result = db.Column(db.Text, nullable=True)
    actual_result = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(20), nullable=False, default='未執行')
    # 舊的 tags 欄位已被下方的 relationship 取代
    # tags = db.Column(db.Text, nullable=True)
    notes = db.Column(db.Text, nullable=True)
    reference = db.Column(db.String(200), nullable=True)

    # 步驟 3: 建立與 Tag 模型的多對多關係
    tags = db.relationship('Tag', secondary=test_case_tags, lazy='subquery',
                           backref=db.backref('test_cases', lazy=True))

# --- 3. 輔助函式 ---

def process_tags(tags_string):
    """
    處理傳入的標籤字串，返回 Tag 物件列表。
    如果標籤不存在，會自動建立。
    """
    if not tags_string or not isinstance(tags_string, str):
        return []
    
    processed_tags = []
    tag_names = [name.strip().lower() for name in tags_string.split(',') if name.strip()]
    
    for name in tag_names:
        tag = Tag.query.filter_by(name=name).first()
        if not tag:
            tag = Tag(name=name)
            db.session.add(tag)
        processed_tags.append(tag)
        
    # 因為可能新增 tag，先 flush 以確保 tag 物件有 id
    db.session.flush()
    return processed_tags

def load_category_rules():
    """從 category_rules.json 檔案載入分類規則。"""
    try:
        rules_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'category_rules.json')
        with open(rules_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        print("警告：'category_rules.json' 檔案找不到或格式錯誤，將無法進行自動分類。")
        return []

CATEGORY_RULES = load_category_rules()

def categorize_case(case_data, product_type):
    """
    根據載入的規則(CATEGORY_RULES)與指定的產品別，
    比對測試案例的內容並回傳最精確的分類。
    """
    rules_for_product = CATEGORY_RULES.get(product_type, [])
    if not rules_for_product:
        return "其他", "未分類"

    text_to_check = (
        f"{case_data.get('測試項目', '')} "
        f"{case_data.get('測試目的', '')} "
        f"{case_data.get('測試步驟', '')} "
        f"{case_data.get('預期結果', '')} "
        f"{case_data.get('category', '')}"
    )
    
    for rule in rules_for_product:
        for keyword in rule['keywords']:
            if keyword in text_to_check:
                return rule['main_category'], rule['sub_category']
            
    return "其他", "未分類"

# --- 4. 路由 (Web 頁面邏輯) ---

@app.context_processor
def inject_status_options():
    status_options = ['未執行', '通過', '失敗', '阻塞']
    return dict(status_options=status_options)

# app.py

@app.route('/dashboard')
def dashboard():
    # --- 1. 狀態統計圓餅圖數據 ---
    status_distribution = db.session.query(
        TestCase.status, 
        func.count(TestCase.status)
    ).group_by(TestCase.status).all()
    
    pie_chart_data = {
        'labels': [status[0] for status in status_distribution],
        'data': [status[1] for status in status_distribution]
    }

    # --- 2. 各主要分類的測試進度數據 ---
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
            'category': category_name,
            'total': total,
            'completed': completed,
            'passed': passed,
            'completion_percentage': round(completion_percentage, 1),
            'pass_percentage': round(pass_percentage, 1)
        })

    # ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
    # ★ 核心修正點：在這裡計算總結數據 (summary_data) ★
    # ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
    summary_data = {
        'total_cases': sum(item['total'] for item in progress_data),
        'completed_cases': sum(item['completed'] for item in progress_data),
        'passed_cases': sum(item['passed'] for item in progress_data)
    }
    # ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

    return render_template(
        'dashboard.html',
        pie_chart_data=json.dumps(pie_chart_data),
        progress_data=progress_data,
        summary_data=summary_data,  # <-- ★ 確保 summary_data 被傳遞給前端 ★
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

    # ★★★ 核心修正點：建立一個更完整的假的 pagination 物件 ★★★
    # 我們補上了 .pages 屬性，並將其設為 0，
    # 這樣 {% if pagination and pagination.pages > 1 %} 的判斷就不會出錯。
    mock_pagination = SimpleNamespace(
        total=len(found_cases),
        pages=0, # <-- ★★★ 補上這一行 ★★★
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
        global_preconditions = CATEGORY_RULES.get('global_preconditions', {})
        global_precondition = global_preconditions.get(selected_product)

    query = TestCase.query
    if selected_product:
        query = query.filter_by(product_type=selected_product)
    if selected_main_category:
        query = query.filter_by(main_category=selected_main_category)
    if selected_sub_category:
        query = query.filter_by(sub_category=selected_sub_category)
            
    pagination = query.order_by(TestCase.case_id).paginate(page=page, per_page=50, error_out=False)
    cases_to_display = pagination.items
    
    return render_template('cases.html', 
                           cases=cases_to_display, 
                           tree_data=tree_data,
                           pagination=pagination,
                           selected_product=selected_product, 
                           selected_main_category=selected_main_category,
                           selected_sub_category=selected_sub_category,
                           global_precondition=global_precondition)

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
        # ★ 標籤處理更新
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
        
        # ★ 標籤處理更新
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
    flash('測試案例已成功刪除。', 'info')
    return redirect(url_for('index'))

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
    # ★ 刪除標籤邏輯更新
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
        'sub_category': request.form.get('sub_category')
    }
    cleaned_redirect_params = {k: v for k, v in redirect_params.items() if v}

    if not case_ids or not new_tag_name:
        flash('未選擇任何案例或未輸入標籤。', 'warning')
        return redirect(url_for('index', **cleaned_redirect_params))

    # ★ 批次新增標籤邏輯更新
    # 先找到或建立 Tag 物件
    tag_to_add = Tag.query.filter_by(name=new_tag_name).first()
    if not tag_to_add:
        tag_to_add = Tag(name=new_tag_name)
        db.session.add(tag_to_add)

    # 遍歷所有選擇的 case
    cases_to_update = TestCase.query.filter(TestCase.id.in_(case_ids)).all()
    for case in cases_to_update:
        if tag_to_add not in case.tags:
            case.tags.append(tag_to_add)
            
    db.session.commit()
    flash(f'已為 {len(case_ids)} 個案例成功新增標籤 "{new_tag_name}"！', 'success')
    return redirect(url_for('index', **cleaned_redirect_params))
    
@app.route('/edit-notes/<int:id>', methods=['GET', 'POST'])
def edit_notes(id):
    case = TestCase.query.get_or_404(id)
    if request.method == 'POST':
        case.notes = request.form.get('notes', '')
        db.session.commit()
        return render_template('partials/_notes_display.html', case=case)
    return render_template('partials/_notes_edit.html', case=case)

@app.route('/display-notes/<int:id>')
def display_notes(id):
    case = TestCase.query.get_or_404(id)
    return render_template('partials/_notes_display.html', case=case)

@app.route('/upload', methods=['GET', 'POST'])
def upload_page():
    # 延後匯入以避免循環依賴
    from services import process_excel_file
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
        
        # ★ 注意：因為 process_tags 會自動 commit，這裡的 commit 邏輯需要調整
        # 改為在 process_excel_file 內部處理 commit
        # if total_imported_count > 0 and not has_error:
        #     db.session.commit()

        if not has_error and total_imported_count > 0:
             flash(f'所有檔案處理完畢！共成功匯入 {total_imported_count} 筆新案例！', 'success')
        elif not has_error and total_imported_count == 0:
            flash('所有檔案處理完畢，但沒有匯入任何新案例 (可能 Case ID 皆已存在)。', 'info')

        return redirect(url_for('index'))

    return render_template('upload.html')

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- 5. 啟動程式 ---
if __name__ == '__main__':
    # 刪除 db.create_all()，由 Flask-Migrate 管理
    # with app.app_context():
    #     db.create_all()
    app.run(debug=True, port=5001)