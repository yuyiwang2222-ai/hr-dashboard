"""
透過 Outlook 寄送人力分析報告
"""

import win32com.client as win32
import os
from datetime import date

def send_report_email():
    # 設定
    today = date.today()
    report_filename = f"人力分析報告_{today.strftime('%Y%m%d')}.pdf"
    script_dir = os.path.dirname(os.path.abspath(__file__))
    report_dir = os.path.join(script_dir, "報告")
    report_path = os.path.join(report_dir, report_filename)
    
    # 檢查報告是否存在
    if not os.path.exists(report_path):
        print(f"❌ 找不到報告：{report_path}")
        print("請先執行 python generate_pdf_report.py 產生報告")
        return False
    
    try:
        # 連接 Outlook
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)  # 0 = MailItem
        
        # 設定收件人
        mail.To = "cheihyi@df-recycle.com.tw"
        
        # 設定主旨
        mail.Subject = f"大豐環保人力分析報告 - {today.strftime('%Y年%m月%d日')}"
        
        # 設定內容
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
大豐環保人力儀表板
"""
        
        # 附加報告
        mail.Attachments.Add(report_path)
        
        # 發送
        mail.Send()
        
        print(f"✅ 郵件已成功寄送至 cheihyi@df-recycle.com.tw")
        print(f"📎 附件：{report_filename}")
        return True
        
    except Exception as e:
        print(f"❌ 寄送失敗：{e}")
        print("\n可能原因：")
        print("1. Outlook 未安裝或未登入")
        print("2. 需要先安裝 pywin32：pip install pywin32")
        return False

if __name__ == "__main__":
    print("📧 正在透過 Outlook 寄送人力分析報告...")
    send_report_email()
