# utils.py
import os
import json
from extensions import db
from models import Tag

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
        return {} # 返回一個空的 dict 更安全

# --- ▼▼▼ 核心修改處 1：新增 update_global_preconditions 函式 ▼▼▼ ---
def update_global_preconditions(product_type, new_preconditions):
    """
    更新 category_rules.json 檔案中的全域前置條件。
    """
    if not new_preconditions or not isinstance(new_preconditions, str):
        return # 如果沒有提供新的條件文字，則不執行任何操作

    rules_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'category_rules.json')
    
    try:
        # 讀取現有的規則
        with open(rules_path, 'r', encoding='utf-8') as f:
            rules_data = json.load(f)
        
        # 如果 'global_preconditions' 鍵不存在，就建立它
        if 'global_preconditions' not in rules_data:
            rules_data['global_preconditions'] = {}
            
        # 更新或新增指定產品類型的前置條件
        rules_data['global_preconditions'][product_type] = new_preconditions.strip()
        
        # 將更新後的內容寫回檔案
        with open(rules_path, 'w', encoding='utf-8') as f:
            json.dump(rules_data, f, ensure_ascii=False, indent=2)
            
    except Exception as e:
        print(f"錯誤：更新全域前置條件失敗 - {e}")
# --- ▲▲▲ 修改結束 ▲▲▲ ---


def categorize_case(case_data, product_type):
    """
    根據載入的規則與指定的產品別，
    比對測試案例的內容並回傳最精確的分類。
    """
    CATEGORY_RULES = load_category_rules()
    
    rules_for_product = CATEGORY_RULES.get(product_type)
    if not rules_for_product:
        return "其他", "未分類"

    text_to_check = (
        f"{case_data.get('測試項目', '')} "
        f"{case_data.get('測試目的', '')} "
        f"{case_data.get('測試步驟', '')} "
        f"{case_data.get('預期結果', '')} "
        f"{case_data.get('category', '')}"
    ).lower()
    
    if isinstance(rules_for_product, list) and rules_for_product and isinstance(rules_for_product[0], str):
        for keyword in rules_for_product:
            if keyword.lower() in text_to_check:
                return product_type, None
        return product_type, '未分類'

    if isinstance(rules_for_product, list) and rules_for_product and isinstance(rules_for_product[0], dict):
        for rule in rules_for_product:
            for keyword in rule['keywords']:
                if keyword.lower() in text_to_check:
                    return rule['main_category'], rule['sub_category']
            
    return "其他", "未分類"