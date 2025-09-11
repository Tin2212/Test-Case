import os
import json
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_sqlalchemy import SQLAlchemy
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

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- 2. 資料庫模型定義 ---
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
    tags = db.Column(db.Text, nullable=True)
    notes = db.Column(db.Text, nullable=True)
    reference = db.Column(db.String(200), nullable=True)

# --- 3. 輔助函式 ---

def load_category_rules():
    """從 category_rules.json 檔案載入分類規則。"""
    try:
        # 使用 os.path.join 確保路徑在不同作業系統上都正確
        rules_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'category_rules.json')
        with open(rules_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        # 如果檔案遺失或格式錯誤，回傳空列表以避免程式崩潰
        print("警告：'category_rules.json' 檔案找不到或格式錯誤，將無法進行自動分類。")
        return []

# 在程式啟動時就載入規則
CATEGORY_RULES = load_category_rules()

def categorize_case(case_data, product_type):
    """
    根據載入的規則(CATEGORY_RULES)與指定的產品別，
    比對測試案例的內容並回傳最精確的分類。
    """
    # 根據產品類型，從大的規則庫中選取對應的規則列表
    rules_for_product = CATEGORY_RULES.get(product_type, [])
    
    # 如果找不到對應產品的規則，直接回傳未分類
    if not rules_for_product:
        return "其他", "未分類"

    # 要進行比對的文字內容，現在包含更豐富的資訊
    text_to_check = (
        f"{case_data.get('測試項目', '')} "
        f"{case_data.get('測試目的', '')} "
        f"{case_data.get('測試步驟', '')} "
        f"{case_data.get('預期結果', '')} "
        f"{case_data.get('category', '')}"  # 保留原始分類作為參考
    )
    
    # 遍歷指定產品的規則列表
    for rule in rules_for_product:
        for keyword in rule['keywords']:
            if keyword in text_to_check:
                return rule['main_category'], rule['sub_category']
            
    # 如果所有規則都沒匹配上，則歸為未分類
    return "其他", "未分類"

# --- 4. 路由 (Web 頁面邏輯) ---

@app.context_processor
def inject_status_options():
    status_options = ['未執行', '通過', '失敗', '阻塞']
    return dict(status_options=status_options)

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

    # --- 核心修改點：讀取並傳遞全域前置條件 ---
    global_precondition = None
    if selected_product:
        # 從載入的 CATEGORY_RULES 中找到 global_preconditions
        global_preconditions = CATEGORY_RULES.get('global_preconditions', {})
        # 根據當前選擇的產品，取得對應的條件文字
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
                           global_precondition=global_precondition) # <-- 傳遞新變數到前端
@app.route('/add', methods=['GET', 'POST'])
def add_case():
    # (此函式內容不變)
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
            tags=case_data.get('tags'), notes=case_data.get('notes'), reference=case_data.get('reference')
        )
        db.session.add(new_case)
        db.session.commit()
        flash('測試案例已成功新增！', 'success')
        return redirect(url_for('index'))
    return render_template('case_form.html', title="新增測試案例", case=None)

@app.route('/edit/<int:id>', methods=['GET', 'POST'])
def edit_case(id):
    # (此函式內容不變)
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
        case_to_edit.tags = case_data.get('tags')
        case_to_edit.notes = case_data.get('notes')
        case_to_edit.reference = case_data.get('reference')
        db.session.commit()
        flash('測試案例已成功更新！', 'success')
        return redirect(url_for('index'))
    return render_template('case_form.html', title="編輯測試案例", case=case_to_edit)

@app.route('/delete/<int:id>', methods=['POST'])
def delete_case(id):
    # (此函式內容不變)
    case_to_delete = TestCase.query.get_or_404(id)
    db.session.delete(case_to_delete)
    db.session.commit()
    flash('測試案例已成功刪除。', 'info')
    return redirect(url_for('index'))

@app.route('/edit-status-result/<int:id>', methods=['GET', 'POST'])
def edit_status_result(id):
    # (此函式內容不變)
    case = TestCase.query.get_or_404(id)
    if request.method == 'POST':
        case.status = request.form.get('status')
        case.actual_result = request.form.get('actual_result', '')
        db.session.commit()
        return render_template('partials/_status_result_display.html', case=case)
    return render_template('partials/_status_result_edit.html', case=case)

@app.route('/display-status-result/<int:id>')
def display_status_result(id):
    # (此函式內容不變)
    case = TestCase.query.get_or_404(id)
    return render_template('partials/_status_result_display.html', case=case)

@app.route('/delete-tag/<int:id>', methods=['POST'])
def delete_tag(id):
    # (此函式內容不變)
    case = TestCase.query.get_or_404(id)
    tag_to_delete = request.form.get('tag')
    if case.tags and tag_to_delete:
        tags_list = [t.strip() for t in case.tags.split(',') if t.strip()]
        if tag_to_delete in tags_list:
            tags_list.remove(tag_to_delete)
        case.tags = ','.join(sorted(tags_list))
        db.session.commit()
    return render_template('partials/_tags_display.html', case=case)

@app.route('/bulk-add-tag', methods=['POST'])
def bulk_add_tag():
    case_ids = request.form.getlist('case_ids')
    new_tag = request.form.get('new_tag', '').strip()

    # 1. 從表單中獲取當前的篩選條件
    product = request.form.get('product')
    main_category = request.form.get('main_category')
    sub_category = request.form.get('sub_category')

    # 2. 準備一個字典，用於建立重新導向的 URL
    #    這裡會過濾掉值為空或 None 的參數
    redirect_params = {
        'product': product,
        'main_category': main_category,
        'sub_category': sub_category
    }
    cleaned_redirect_params = {k: v for k, v in redirect_params.items() if v}

    if not case_ids or not new_tag:
        flash('未選擇任何案例或未輸入標籤。', 'warning')
        # 3. 使用清理過的參數進行重新導向
        return redirect(url_for('index', **cleaned_redirect_params))

    for case_id in case_ids:
        case = TestCase.query.get(case_id)
        if case:
            existing_tags = set(t.strip() for t in case.tags.split(',') if t.strip())
            existing_tags.add(new_tag)
            case.tags = ','.join(sorted(list(existing_tags)))
            
    db.session.commit()
    flash(f'已為 {len(case_ids)} 個案例成功新增標籤 "{new_tag}"！', 'success')

    # 4. 在成功時，也使用清理過的參數進行重新導向
    return redirect(url_for('index', **cleaned_redirect_params))
    
@app.route('/edit-notes/<int:id>', methods=['GET', 'POST'])
def edit_notes(id):
    """(HTMX) 處理備註欄位的即時編輯"""
    case = TestCase.query.get_or_404(id)
    if request.method == 'POST':
        case.notes = request.form.get('notes', '')
        db.session.commit()
        return render_template('partials/_notes_display.html', case=case)
    return render_template('partials/_notes_edit.html', case=case)

@app.route('/display-notes/<int:id>')
def display_notes(id):
    """(HTMX) 專門給"取消"按鈕使用，只回傳備註的顯示內容片段"""
    case = TestCase.query.get_or_404(id)
    return render_template('partials/_notes_display.html', case=case)

@app.route('/upload', methods=['GET', 'POST'])
def upload_page():
    # (此函式內容不變)
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
                product_type = '郵件閘道' if filename.lower().startswith('gateway') else '郵件歸檔'
                try:
                    xls = pd.ExcelFile(file.stream)
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name).fillna('')
                        for index, row in df.iterrows():
                            if not row.get('Case ID'): continue
                            exists = TestCase.query.filter_by(case_id=row['Case ID']).first()
                            if not exists:
                                case_data_for_cat = row.to_dict()
                                case_data_for_cat['category'] = sheet_name
                                main_cat, sub_cat = categorize_case(case_data_for_cat, product_type)
                                new_case = TestCase(
                                    product_type=product_type, category=sheet_name,
                                    main_category=main_cat, sub_category=sub_cat,
                                    case_id=row.get('Case ID', ''), test_item=row.get('測試項目', ''),
                                    test_purpose=row.get('測試目的', ''), preconditions=row.get('前置條件', ''),
                                    test_steps=row.get('測試步驟', ''), expected_result=row.get('預期結果', ''),
                                    actual_result=row.get('實際結果', ''), status=row.get('狀態', '未執行'),
                                    tags=row.get('標籤', ''), notes=row.get('備註', ''),
                                    reference=row.get('參考資料', '')
                                )
                                db.session.add(new_case)
                                total_imported_count += 1
                except Exception as e:
                    has_error = True
                    flash(f'處理檔案 "{filename}" 時發生錯誤：{e}', 'danger')
                    db.session.rollback()
        if total_imported_count > 0:
            db.session.commit()
        if not has_error and total_imported_count > 0:
             flash(f'所有檔案處理完畢！共成功匯入 {total_imported_count} 筆新案例！', 'success')
        elif not has_error and total_imported_count == 0:
            flash('所有檔案處理完畢，但沒有匯入任何新案例(可能 Case ID 皆已存在)。', 'info')
        return redirect(url_for('index'))
    return render_template('upload.html')

# --- 5. 啟動程式 ---
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True, port=5001)