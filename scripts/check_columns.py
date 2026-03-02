"""
大豐環保人力儀表板 - 欄位檢視工具
============================================
檢視「數據資料夾」中資料檔案的欄位名稱
用於設定 config.yaml 中的欄位對應

使用方式：
  python scripts/check_columns.py
"""

import os
import sys
import glob
import pandas as pd

# 設定路徑
script_dir = os.path.dirname(os.path.abspath(__file__))
project_dir = os.path.dirname(script_dir)
os.chdir(project_dir)

# 外部資料夾路徑
EXTERNAL_PATH = "./數據資料夾"


def main():
    print("=" * 60)
    print("📊 大豐環保人力儀表板 - 欄位檢視工具")
    print("=" * 60)
    print()
    
    if not os.path.exists(EXTERNAL_PATH):
        print(f"❌ 資料夾不存在：{EXTERNAL_PATH}")
        return
    
    # 尋找所有 Excel 檔案
    xlsx_files = glob.glob(os.path.join(EXTERNAL_PATH, "*.xlsx"))
    xls_files = glob.glob(os.path.join(EXTERNAL_PATH, "*.xls"))
    all_files = xlsx_files + [f for f in xls_files if not f.endswith('.xlsx')]
    
    if not all_files:
        print(f"⚠️ 資料夾中無 Excel 檔案")
        print(f"   請將員工資料檔案放入：{os.path.abspath(EXTERNAL_PATH)}")
        return
    
    print(f"📂 找到 {len(all_files)} 個檔案：")
    print("-" * 60)
    
    for file_path in all_files:
        filename = os.path.basename(file_path)
        mod_time = os.path.getmtime(file_path)
        mod_date = pd.Timestamp.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M')
        
        print(f"\n📄 {filename}")
        print(f"   修改時間：{mod_date}")
        
        try:
            # 讀取檔案
            if file_path.endswith('.xls') and not file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path, sheet_name=0, engine='xlrd')
            else:
                df = pd.read_excel(file_path, sheet_name=0)
            
            print(f"   資料筆數：{len(df)}")
            print(f"   欄位數量：{len(df.columns)}")
            print()
            print("   📋 欄位名稱：")
            print("   " + "-" * 50)
            
            for i, col in enumerate(df.columns, 1):
                # 顯示欄位名稱和前幾筆資料範例
                sample = df[col].dropna().head(3).tolist()
                sample_str = str(sample)[:40] + "..." if len(str(sample)) > 40 else str(sample)
                print(f"   {i:2d}. {col:<20} 範例: {sample_str}")
            
        except Exception as e:
            print(f"   ❌ 讀取錯誤：{e}")
    
    print()
    print("-" * 60)
    print()
    print("💡 請對照上述欄位名稱，編輯 config.yaml 中的 column_mapping 設定")
    print()
    print("   範例：")
    print("   column_mapping:")
    print("     employee_id: \"工號\"      # 你的欄位名稱")
    print("     name: \"姓名\"")
    print("     department: \"部門\"")
    print("     position: \"職稱\"")
    print("     hire_date: \"到職日\"")
    print("     leave_date: \"離職日\"")
    print()


if __name__ == "__main__":
    main()
