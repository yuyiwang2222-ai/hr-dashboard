"""
大豐環保人力分析報告 - 自動化腳本
============================================
自動抓取資料 → 產生 PDF → 寄送郵件

使用方式：
1. 手動執行：python auto_report.py
2. 排程執行：透過 Windows 工作排程器
"""

import os
import sys
from datetime import date, datetime

# 設定工作目錄
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)
sys.path.insert(0, script_dir)

def log(message):
    """記錄訊息"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{timestamp}] {message}")
    
    # 同時寫入日誌檔
    log_dir = os.path.join(script_dir, "報告")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "auto_report.log")
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"[{timestamp}] {message}\n")


def sync_and_push_to_github():
    """
    同步資料並推送到 GitHub，自動更新 Streamlit Cloud
    """
    import subprocess
    
    try:
        # 執行 sync_data.py 同步資料
        log("正在同步資料...")
        sync_script = os.path.join(script_dir, "sync_data.py")
        result = subprocess.run(
            [sys.executable, sync_script],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            cwd=script_dir
        )
        
        if result.returncode != 0:
            log(f"同步資料失敗：{result.stderr}")
            return False
        
        # Git add, commit, push
        log("正在推送到 GitHub...")
        
        # git add
        subprocess.run(["git", "add", "."], cwd=script_dir, check=True)
        
        # git commit
        today = date.today().strftime('%Y-%m-%d')
        commit_result = subprocess.run(
            ["git", "commit", "-m", f"自動更新資料 {today}"],
            cwd=script_dir,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )
        
        # 如果沒有變更，commit 會失敗，但這不是錯誤
        if commit_result.stdout and "nothing to commit" in commit_result.stdout:
            log("沒有新的變更需要提交")
            return True
        if commit_result.stderr and "nothing to commit" in commit_result.stderr:
            log("沒有新的變更需要提交")
            return True
        
        # git push
        push_result = subprocess.run(
            ["git", "push"],
            cwd=script_dir,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            check=True
        )
        
        log("Git 推送成功，Streamlit Cloud 將自動更新")
        return True
        
    except subprocess.CalledProcessError as e:
        log(f"Git 操作失敗：{e}")
        return False
    except Exception as e:
        log(f"同步錯誤：{str(e)}")
        return False


def run_auto_report(send_email=True, recipients=None):
    """
    執行自動報告流程
    
    Args:
        send_email: 是否寄送郵件（預設 True）
        recipients: 收件人列表（預設使用設定檔）
    """
    log("=" * 50)
    log("開始執行自動報告流程")
    
    # 預設收件人
    if recipients is None:
        recipients = ["chiehyi@df-recycle.com.tw"]
    
    try:
        # Step 1: 同步資料並推送到 GitHub
        log("步驟 1/4：同步資料到 GitHub...")
        git_success = sync_and_push_to_github()
        if git_success:
            log("✅ 資料已同步並推送到 GitHub")
        else:
            log("⚠️ Git 推送失敗，繼續執行報告生成")
        
        # Step 2: 清除快取，確保抓取最新資料
        log("步驟 2/4：清除快取...")
        from src.data_loader import clear_cache
        clear_cache()
        log("✅ 快取已清除")
        
        # Step 3: 產生 PDF 報告
        log("步驟 3/4：產生 PDF 報告...")
        from generate_pdf_report import generate_pdf_report
        report_path = generate_pdf_report()
        
        if report_path and os.path.exists(report_path):
            log(f"✅ 報告已生成：{report_path}")
        else:
            log("❌ 報告生成失敗")
            return False
        
        # Step 4: 寄送郵件
        if send_email:
            log("步驟 4/4：寄送郵件...")
            success = send_report_with_outlook(report_path, recipients)
            if success:
                log(f"✅ 郵件已寄送至：{', '.join(recipients)}")
            else:
                log("❌ 郵件寄送失敗")
                return False
        else:
            log("步驟 4/4：跳過寄送郵件")
        
        log("🎉 自動報告流程完成！")
        return True
        
    except Exception as e:
        log(f"❌ 執行錯誤：{str(e)}")
        import traceback
        log(traceback.format_exc())
        return False


def send_report_with_outlook(report_path, recipients):
    """透過 Outlook 寄送報告"""
    try:
        import win32com.client as win32
        
        today = date.today()
        
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        
        # 設定收件人（支援多人）
        mail.To = ";".join(recipients)
        
        mail.Subject = f"大豐環保人力分析報告 - {today.strftime('%Y年%m月%d日')}"
        
        mail.Body = f"""您好，

附件為 {today.strftime('%Y年%m月%d日')} 的人力分析報告。

報告內容包含：
📊 人力概況摘要
📊 各部門人數比較（當月 vs 2025年1月）
📊 各事業部非一線人員比較
📊 各事業部行政幕僚人員比較
📋 即將離職人員名單

如有任何問題，請與人資部聯繫。

---
大豐環保人力儀表板（自動發送）
"""
        
        mail.Attachments.Add(report_path)
        mail.Send()
        
        return True
        
    except Exception as e:
        log(f"Outlook 錯誤：{str(e)}")
        return False


# ============================================
# 排程設定
# ============================================
"""
【Windows 工作排程器設定步驟】

1. 開啟「工作排程器」
   - Win + R → taskschd.msc → Enter

2. 建立基本工作
   - 右側點選「建立基本工作」
   - 名稱：大豐環保人力報告
   - 說明：每日自動產生人力分析報告並寄送

3. 觸發程序（選擇執行時機）
   選項 A：每日
   - 選擇「每天」
   - 設定時間（例如：08:00）
   
   選項 B：每週一
   - 選擇「每週」
   - 勾選「星期一」
   - 設定時間（例如：08:00）

4. 動作
   - 選擇「啟動程式」
   - 程式：C:\\Users\\chiehyi\\AppData\\Local\\Programs\\Python\\Python313\\python.exe
   - 引數：auto_report.py
   - 起始位置：C:\\Users\\chiehyi\\OneDrive - 大豐環保科技股份有限公司\\文件\\Claude\\Agents-人力變化

5. 完成
   - 勾選「當我按完成時，開啟這個工作的內容對話方塊」
   - 在「一般」頁籤：勾選「以最高權限執行」
   - 在「條件」頁籤：取消勾選「只有在電腦使用 AC 電源時才啟動這個工作」

【注意事項】
- 電腦必須開機且已登入
- Outlook 必須已設定好郵件帳戶
- 如果使用 VPN，排程執行時 VPN 需連線
"""


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='大豐環保人力分析報告自動化')
    parser.add_argument('--no-email', action='store_true', help='只產生報告，不寄送郵件')
    parser.add_argument('--to', nargs='+', help='收件人（可多人，空格分隔）')
    
    args = parser.parse_args()
    
    recipients = args.to if args.to else None
    send_email = not args.no_email
    
    success = run_auto_report(send_email=send_email, recipients=recipients)
    sys.exit(0 if success else 1)
