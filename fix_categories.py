# fix_categories.py
import json
from app import app, db, TestCase, categorize_case

def run_fix():
    """
    重新執行所有測試案例的分類邏輯，並更新資料庫。
    """
    # 確保在 Flask 應用程式的上下文中執行
    with app.app_context():
        all_cases = TestCase.query.all()
        
        if not all_cases:
            print("資料庫中沒有找到任何測試案例。")
            return

        print(f"開始修正 {len(all_cases)} 筆測試案例的分類...")
        
        updated_count = 0
        for case in all_cases:
            # 建立一個模擬的 case_data 字典給 categorize_case 函式使用
            case_data = {
                '測試項目': case.test_item,
                '測試步驟': case.test_steps,
                'category': case.category  # 使用原始的 Excel 工作表名稱
            }
            
            # 使用修正後的邏輯重新獲取分類
            new_main_cat, new_sub_cat = categorize_case(case_data)
            
            # 只有在分類有變動時才更新
            if case.main_category != new_main_cat or case.sub_category != new_sub_cat:
                print(f"  - [修正] Case ID: {case.case_id}")
                print(f"    舊分類: {case.main_category} / {case.sub_category}")
                print(f"    新分類: {new_main_cat} / {new_sub_cat}")
                case.main_category = new_main_cat
                case.sub_category = new_sub_cat
                updated_count += 1

        if updated_count > 0:
            # 將所有變動一次性提交到資料庫
            db.session.commit()
            print(f"\n修正完成！共更新了 {updated_count} 筆案例的分類。")
        else:
            print("\n檢查完成，所有案例的分類都已是最新，無需更新。")

if __name__ == '__main__':
    run_fix()