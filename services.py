import pandas as pd
import re
from sqlalchemy.exc import IntegrityError
from models import TestCase, Tag
from extensions import db
from utils import categorize_case, process_tags, update_global_preconditions

def process_excel_file(file_stream, filename, selected_product_type):
    """
    處理上傳的 Excel 檔案，並將測試案例存入資料庫。
    """
    try:
        all_sheets_dict = pd.read_excel(file_stream, engine='openpyxl', sheet_name=None, header=None)
        
        processed_sheets_data = []
        is_first_sheet = True
        
        # --- ▼▼▼ 核心修改處 1：在解析檔名後儲存 precond_key ▼▼▼ ---
        precondition_key = selected_product_type # 預設使用產品類型作為 key

        # 遍歷每一個讀取進來的工作表
        for sheet_name, sheet_df in all_sheets_dict.items():
            if sheet_df.empty:
                continue

            if is_first_sheet:
                # 只有 Spec 和 Tests 類型需要從檔名決定 precond_key
                if selected_product_type in ["Smail-Spec", "Smail-Tests"]:
                    match = re.search(r'(Spec|Tests)[#-_]?(\d{3,})', filename, re.IGNORECASE)
                    if match:
                        # 如果檔名匹配成功，就用 Spec#ID 作為儲存前置條件的 key
                        precondition_key = f"{match.group(1).capitalize()}#{match.group(2)}"

                if pd.notna(sheet_df.iloc[0, 0]):
                    preconditions_text = str(sheet_df.iloc[0, 0])
                    # 使用解析出的 precondition_key 來更新 JSON 檔案
                    update_global_preconditions(precondition_key, preconditions_text)
                
                is_first_sheet = False
        # --- ▲▲▲ 修改結束 ▲▲▲ ---

            header_row_index = -1
            for i, row_series in sheet_df.iterrows():
                for cell_value in row_series:
                    if isinstance(cell_value, str) and 'Case ID' in cell_value:
                        header_row_index = i
                        break
                if header_row_index != -1:
                    break
            
            if header_row_index != -1:
                new_header = sheet_df.iloc[header_row_index]
                case_data_df = sheet_df.iloc[header_row_index + 1:]
                case_data_df.columns = new_header
                processed_sheets_data.append(case_data_df)

        if not processed_sheets_data:
             raise ValueError("在 Excel 檔案中找不到任何包含 'Case ID' 標頭的工作表。")

        df = pd.concat(processed_sheets_data, ignore_index=True)
        df.columns = df.columns.str.strip()
    except Exception as e:
        raise ValueError(f"無法讀取或解析 Excel 檔案：{e}")

    df = df.fillna('')

    required_columns = ['Case ID', '測試項目']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Excel 檔案中缺少必要的欄位：'{col}'")

    imported_count = 0
    existing_case_ids = {case.case_id for case in TestCase.query.with_entities(TestCase.case_id).all()}

    for index, row in df.iterrows():
        case_id = str(row.get('Case ID', '')).strip()
        if not case_id or case_id in existing_case_ids:
            continue

        case_data = row.to_dict()
        product_type = selected_product_type
        main_cat = None
        sub_cat = None

        if product_type in ["Smail-Spec", "Smail-Tests"]:
            product_type = selected_product_type
            match = re.search(r'(Spec|Tests)[#-_]?(\d{3,})', filename, re.IGNORECASE)
            if match:
                main_cat = f"{match.group(1).capitalize()}#{match.group(2)}"
            else:
                main_cat = '未分類'
            sub_cat = None
        else:
            main_cat, sub_cat = categorize_case(case_data, product_type)

        new_case = TestCase(
            case_id=case_id,
            product_type=product_type,
            category=str(row.get('category', '')),
            main_category=main_cat,
            sub_category=sub_cat,
            test_item=str(row.get('測試項目', '')),
            test_purpose=str(row.get('測試目的', '')),
            preconditions=str(row.get('前置條件', '')),
            test_steps=str(row.get('測試步驟', '')),
            expected_result=str(row.get('預期結果', '')),
            status='未執行',
            notes=str(row.get('備註', '')),
            reference=str(row.get('參考資料', '')),
            tags=process_tags(str(row.get('標籤', '')))
        )
        db.session.add(new_case)
        existing_case_ids.add(case_id)
        imported_count += 1

    if imported_count > 0:
        try:
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            raise ValueError("儲存資料時發生錯誤，可能存在重複的 Case ID 或欄位不符。")
        except Exception as e:
            db.session.rollback()
            raise IOError(f"寫入資料庫時發生未知錯誤：{e}")

    return imported_count