import os
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
def categorize_case(case_data):
    """根據測試案例的內容(測試項目/步驟/原始分類)，比對手冊結構，回傳最精確的主分類和子分類。"""
    category_str = case_data.get('category', '')
    text_to_check = f"{case_data.get('測試項目', '')} {case_data.get('測試步驟', '')} {category_str}"
    
    KEYWORD_MAP = {
        ("登入", "登出", "使用者登入與介面"): ("使用者介面", "登入與登出"),
        ("收件匣", "寄件備份", "垃圾郵件", "郵件資料夾"): ("使用者介面", "郵件資料夾"),
        ("搜尋郵件",): ("使用者介面", "搜尋郵件"),
        ("偏好設定", "使用者個人介面與偏好設定"): ("使用者介面", "偏好設定"),
        ("事件紀錄",): ("使用者介面", "事件紀錄"),
        ("參數設定",): ("系統", "參數設定"),
        ("紀錄檢視", "使用紀錄", "查看紀錄", "紀錄與維運"): ("系統", "紀錄檢視"),
        ("系統資訊",): ("系統", "系統資訊"),
        ("授權資訊",): ("系統", "授權資訊"),
        ("系統更新",): ("系統", "系統更新"),
        ("組態鎖定",): ("系統", "組態鎖定"),
        ("特殊郵件", "索引錯誤", "解析錯誤"): ("系統", "特殊郵件處理"),
        ("網域管理", "新增網域", "刪除網域", "網域與架構功能"): ("網域", "網域管理"),
        ("別名管理", "網域別名"): ("網域", "別名管理"),
        ("閘道架構", "中繼路由", "SMTP 認證"): ("架構", "閘道架構"),
        ("帳號列表", "新增使用者", "帳號列表與環境"): ("帳號", "帳號列表與環境"),
        ("帳號環境",): ("帳號", "帳號列表與環境"),
        ("帳號原則", "密碼原則", "權限與原則"): ("帳號", "帳號原則設定"),
        ("帳號保護", "鎖定帳號"): ("帳號", "帳號原則設定"),
        ("帳號同步", "管理與同步"): ("帳號", "帳號同步"),
        ("帳號認證",): ("帳號", "帳號認證"),
        ("管理者帳號", "管理功能權限"): ("帳號", "管理者帳號"),
        ("群組管理", "新增群組", "群組與郵件查詢"): ("群組", "群組管理"),
        ("郵件總管",): ("郵件", "郵件總管"),
        ("郵件查詢", "進階查詢"): ("郵件", "郵件查詢"),
        ("暫存檔案", "審查與暫存"): ("郵件", "暫存檔案"),
        ("郵件審查", "審查人員"): ("稽核", "審查管理"),
        ("連線封鎖",): ("過濾", "連線封鎖"),("收件人有效性",): ("過濾", "收件人有效性"),
        ("來源檢查", "RBL", "SPF", "DKIM", "DMARC"): ("過濾", "來源檢查"),
        ("允許及封鎖名單",): ("過濾", "允許及封鎖名單"),("垃圾郵件特徵", "URIBL"): ("過濾", "垃圾郵件特徵"),
        ("資料特徵",): ("過濾", "資料特徵設定"),("垃圾郵件通知",): ("過濾", "垃圾郵件通知"),
        ("郵件防毒", "病毒郵件"): ("威脅", "郵件防毒"),("DoS 防禦",): ("威脅", "DoS 防禦"),
        ("威脅郵件特徵", "釣魚網址", "退信攻擊"): ("威脅", "威脅郵件特徵"),("大量發信偵測",): ("威脅", "大量發信偵測"),
        ("報表精靈", "報表紀錄", "自訂樣板", "樣板與紀錄"): ("報表", "報表精靈"),
        ("週期設定", "儲存資源", "生命週期", "設定與紀錄"): ("封存", "週期設定"),
        ("封存紀錄",): ("封存", "封存紀錄"),
        ("功能組合", "交互影響", "整合與生命週期測試"): ("進階測試", "功能組合與整合"),
        ("強健性", "安全性與邊界測試", "負向測試"): ("進階測試", "強健性與安全"),
        ("API與相容性測試", "系統整合"): ("進階測試", "API與整合"),
        ("系統維運", "進階參數與擴展測試"): ("進階測試", "維運與擴展"),
    }
    
    for keywords, (main_cat, sub_cat) in KEYWORD_MAP.items():
        for keyword in keywords:
            if keyword in text_to_check:
                return main_cat, sub_cat
            
    return "其他", "未分類"

def allowed_file(filename):
    """檢查副檔名是否為允許的 .xlsx"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- 4. 路由 (Web 頁面邏輯) ---

@app.context_processor
def inject_status_options():
    status_options = ['未執行', '通過', '失敗', '阻塞']
    return dict(status_options=status_options)

@app.route('/')
def index():
    # (此函式內容不變)
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

    query = TestCase.query
    if selected_product:
        query = query.filter_by(product_type=selected_product)
    if selected_main_category:
        query = query.filter_by(main_category=selected_main_category)
    if selected_sub_category:
        query = query.filter_by(sub_category=selected_sub_category)
            
    cases_to_display = query.order_by(TestCase.case_id).all()
    
    return render_template('cases.html', cases=cases_to_display, tree_data=tree_data,
                           selected_product=selected_product, selected_main_category=selected_main_category,
                           selected_sub_category=selected_sub_category)

@app.route('/add', methods=['GET', 'POST'])
def add_case():
    # (此函式內容不變)
    if request.method == 'POST':
        case_data = request.form.to_dict()
        main_cat, sub_cat = categorize_case(case_data)
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
        main_cat, sub_cat = categorize_case(case_data)
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
    # (此函式內容不變)
    case_ids = request.form.getlist('case_ids')
    new_tag = request.form.get('new_tag', '').strip()
    if not case_ids or not new_tag:
        flash('未選擇任何案例或未輸入標籤。', 'warning')
        return redirect(url_for('index'))
    for case_id in case_ids:
        case = TestCase.query.get(case_id)
        if case:
            existing_tags = set(t.strip() for t in case.tags.split(',') if t.strip())
            existing_tags.add(new_tag)
            case.tags = ','.join(sorted(list(existing_tags)))
    db.session.commit()
    flash(f'已為 {len(case_ids)} 個案例成功新增標籤 "{new_tag}"！', 'success')
    return redirect(url_for('index'))
    
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
                                main_cat, sub_cat = categorize_case(case_data_for_cat)
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