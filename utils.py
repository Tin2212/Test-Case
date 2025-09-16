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

def categorize_case(case_data, product_type):
    """
    根據載入的規則與指定的產品別，
    比對測試案例的內容並回傳最精確的分類。
    """
    # ★★★ 核心修正點 1: 在函式內部呼叫，確保每次都讀取最新檔案 ★★★
    CATEGORY_RULES = load_category_rules()
    
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