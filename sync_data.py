"""
大豐環保人力儀表板 - 資料同步腳本
============================================
將「數據資料夾」的最新資料同步到「data/」資料夾
用於更新 GitHub 倉庫以同步到 Streamlit Cloud

使用方式：
  python sync_data.py
  
執行後需要：
  git add data/
  git commit -m "更新員工資料"
  git push
"""

import os
import shutil
from datetime import datetime

# 來源與目標路徑
SOURCE_FOLDER = "./數據資料夾"
TARGET_FOLDER = "./data"

# 要同步的檔案對應（來源 -> 目標）
FILE_MAPPING = {
    "員工人數.xlsx": "employees.xlsx",
    "事業部對應表.xlsx": "departments.xlsx",
    "每日出勤總表.xlsx": "attendance.xlsx",
}


def sync_files():
    """同步資料檔案"""
    print("=" * 50)
    print("大豐環保人力儀表板 - 資料同步")
    print(f"執行時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)
    
    # 確保目標資料夾存在
    os.makedirs(TARGET_FOLDER, exist_ok=True)
    
    success_count = 0
    
    for source_name, target_name in FILE_MAPPING.items():
        source_path = os.path.join(SOURCE_FOLDER, source_name)
        target_path = os.path.join(TARGET_FOLDER, target_name)
        
        if os.path.exists(source_path):
            # 複製檔案
            shutil.copy2(source_path, target_path)
            
            # 取得檔案資訊
            mod_time = datetime.fromtimestamp(os.path.getmtime(source_path))
            file_size = os.path.getsize(source_path) / 1024  # KB
            
            print(f"✅ {source_name} -> {target_name}")
            print(f"   最後修改：{mod_time.strftime('%Y-%m-%d %H:%M')}")
            print(f"   檔案大小：{file_size:.1f} KB")
            success_count += 1
        else:
            print(f"⚠️ 找不到：{source_path}")
    
    print("-" * 50)
    print(f"同步完成：{success_count}/{len(FILE_MAPPING)} 個檔案")
    
    if success_count > 0:
        print("\n📤 請執行以下指令上傳到 GitHub：")
        print("   git add data/")
        print('   git commit -m "更新員工資料"')
        print("   git push")
    
    return success_count


if __name__ == "__main__":
    sync_files()
