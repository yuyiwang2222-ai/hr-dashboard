"""
大豐環保人力儀表板 - 每週排程任務
============================================
此腳本由 Windows Task Scheduler 每週一 09:00 執行
功能：
1. 產生本週人力快照
2. 檢查異常警報
3. 寄送 Email 週報

使用方式：
  python scripts/weekly_job.py

排程設定（Windows Task Scheduler）：
  - 觸發條件：每週一 09:00
  - 動作：python scripts/weekly_job.py
  - 起始目錄：專案根目錄
"""

import os
import sys
from datetime import date, datetime

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr.encoding != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# 設定工作目錄為專案根目錄
script_dir = os.path.dirname(os.path.abspath(__file__))
project_dir = os.path.dirname(script_dir)
os.chdir(project_dir)
sys.path.insert(0, project_dir)

from src.data_loader import load_config, load_employees, load_departments, save_snapshot
from src.metrics import generate_snapshot_data
from src.alerts import check_all_alerts, get_alerts_summary
from src.email_report import send_weekly_report


def log(message: str):
    """記錄訊息"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{timestamp}] {message}")


def main():
    """主程式"""
    log("=" * 50)
    log("大豐環保人力儀表板 - 每週任務開始")
    log("=" * 50)
    
    # 載入設定與資料
    log("載入設定檔...")
    config = load_config()
    
    log("載入員工資料...")
    employees_df = load_employees()
    
    if len(employees_df) == 0:
        log("❌ 錯誤：找不到員工資料！")
        return 1
    
    log(f"  共 {len(employees_df)} 筆資料")
    
    log("載入部門資料...")
    departments_df = load_departments()
    
    # 步驟 1：產生快照
    log("-" * 50)
    log("步驟 1：產生人力快照")
    
    today = date.today()
    week_number = today.isocalendar()[1]
    snapshot_filename = f"{today.year}-W{week_number:02d}.xlsx"
    snapshot_path = os.path.join(
        config.get('system', {}).get('snapshot_path', 'data/snapshots'),
        snapshot_filename
    )
    
    snapshot_data = generate_snapshot_data(employees_df)
    
    if save_snapshot(snapshot_data, snapshot_path):
        log(f"✅ 快照已儲存：{snapshot_path}")
    else:
        log(f"⚠️ 快照儲存失敗")
    
    # 顯示快照摘要
    log(f"  - 總人數：{snapshot_data['total_headcount']}")
    log(f"  - 本週入職：{snapshot_data['weekly_hires']}")
    log(f"  - 本週離職：{snapshot_data['weekly_leaves']}")
    log(f"  - 月離職率：{snapshot_data['monthly_turnover_rate']:.1f}%")
    
    # 步驟 2：檢查警報
    log("-" * 50)
    log("步驟 2：檢查異常警報")
    
    alerts = check_all_alerts(employees_df, config, departments_df)
    
    if alerts:
        log(f"⚠️ {get_alerts_summary(alerts)}")
        for alert in alerts:
            log(f"  - {alert.title}: {alert.message}")
    else:
        log("✅ 無異常警報")
    
    # 步驟 3：寄送週報
    log("-" * 50)
    log("步驟 3：寄送 Email 週報")
    
    email_enabled = config.get('email', {}).get('enabled', False)
    
    if email_enabled:
        success = send_weekly_report(employees_df, departments_df, config)
        if success:
            log("✅ 週報已成功寄送")
        else:
            log("⚠️ 週報寄送失敗")
    else:
        log("ℹ️ Email 功能已停用，跳過寄送")
    
    # 完成
    log("-" * 50)
    log("✅ 每週任務完成")
    log("=" * 50)
    
    return 0


if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except Exception as e:
        log(f"❌ 執行錯誤：{e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
