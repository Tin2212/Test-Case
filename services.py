# services.py
from app import db, TestCase, categorize_case, process_tags
import pandas as pd

def process_excel_file(file_stream, filename):
    """
    處理單一上傳的 Excel 檔案並匯入測試案例。
    回傳成功匯入的數量。
    """
    imported_count = 0
    product_type = '郵件閘道' if filename.lower().startswith('gateway') else '郵件歸檔'

    xls = pd.ExcelFile(file_stream)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name).fillna('')
        for index, row in df.iterrows():
            if not row.get('Case ID'):
                continue

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
                    notes=row.get('備註', ''),
                    reference=row.get('參考資料', '')
                )

                # ★ 標籤處理更新
                tags_string = row.get('標籤', '')
                new_case.tags = process_tags(tags_string)
                
                db.session.add(new_case)
                imported_count += 1
    
    # 將本次檔案的所有變更一次性提交
    if imported_count > 0:
        db.session.commit()
        
    return imported_count