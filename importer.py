import pandas as pd
import json

def import_and_categorize_excel(filepath):
    """
    讀取一個 Excel 檔案，並將其內容按工作表名稱分類。
    
    Args:
        filepath (str): Excel 檔案的路徑。
        
    Returns:
        dict: 一個巢狀字典，key 是分類名稱，value 是該分類的測試案例列表。
    """
    try:
        xls = pd.ExcelFile(filepath)
        sheet_names = xls.sheet_names
        
        print(f"成功讀取檔案，找到 {len(sheet_names)} 個工作表(分類)。")
        
        categorized_cases = {}
        for sheet_name in sheet_names:
            # 將工作表讀取為 DataFrame
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # 將 DataFrame 轉換為字典列表 (records 格式)
            # fillna('') 確保空的儲存格是空字串，而不是 NaN
            test_cases_list = df.fillna('').to_dict('records')
            
            # 以工作表名稱為 key 存入結果
            categorized_cases[sheet_name] = test_cases_list
            
        return categorized_cases

    except FileNotFoundError:
        print(f"錯誤：找不到檔案 '{filepath}'。")
        return None
    except Exception as e:
        print(f"讀取或處理檔案時發生錯誤：{e}")
        return None

# --- 主程式執行區 ---
if __name__ == "__main__":
    excel_file_path = 'TestPlan.xlsx' # 確保這個檔案和你的 python 腳本在同一個目錄
    
    # 執行導入與分類
    final_data = import_and_categorize_excel(excel_file_path)
    
    # 如果成功，就將結果漂亮地印出來看看
    if final_data:
        print("\n--- 資料匯入與分類結果 ---")
        # 使用 json.dumps 讓巢狀字典的輸出更易讀
        print(json.dumps(final_data, indent=2, ensure_ascii=False))