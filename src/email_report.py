"""
大豐環保人力儀表板 - Email 週報模組
============================================
負責產生並寄送 Email 週報
"""

import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import date, datetime
from typing import Dict, List, Any, Optional

from .metrics import (
    get_current_headcount,
    get_weekly_changes,
    get_monthly_changes,
    get_turnover_rate,
    get_department_stats
)
from .alerts import Alert, check_all_alerts, format_alert_message, get_alerts_summary


def generate_report_content(
    employees_df,
    departments_df,
    config: Dict[str, Any],
    alerts: List[Alert]
) -> str:
    """
    產生週報內容（HTML 格式）
    
    Args:
        employees_df: 員工資料 DataFrame
        departments_df: 部門對照表
        config: 設定字典
        alerts: 警報列表
        
    Returns:
        HTML 格式的週報內容
    """
    today = date.today()
    
    # 計算各項指標
    total_headcount = get_current_headcount(employees_df)
    weekly_changes = get_weekly_changes(employees_df)
    monthly_turnover = get_turnover_rate(employees_df, 'monthly')
    dept_stats = get_department_stats(employees_df)
    
    # 警報摘要
    alerts_summary = get_alerts_summary(alerts)
    alerts_html = ""
    if alerts:
        alerts_html = "<h3>⚠️ 異常警報</h3><ul>"
        for alert in alerts:
            color = '#F44336' if alert.level.value == 'critical' else '#FFC107'
            alerts_html += f'<li style="color: {color};">{alert.title}: {alert.message}</li>'
        alerts_html += "</ul>"
    
    # 部門統計表格
    dept_table_rows = ""
    for _, row in dept_stats.iterrows():
        dept_code = row['department']
        dept_name = dept_code
        # 取得部門名稱
        match = departments_df[departments_df['code'] == dept_code]
        if len(match) > 0:
            dept_name = match.iloc[0]['name']
        
        dept_table_rows += f"""
        <tr>
            <td>{dept_name}</td>
            <td style="text-align: center;">{row['count']}</td>
            <td style="text-align: center;">{row['direct_count']}</td>
            <td style="text-align: center;">{row['indirect_count']}</td>
        </tr>
        """
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{
                font-family: 'Microsoft JhengHei', Arial, sans-serif;
                line-height: 1.6;
                color: #333;
                max-width: 800px;
                margin: 0 auto;
                padding: 20px;
            }}
            .header {{
                background: linear-gradient(135deg, #1E88E5, #42A5F5);
                color: white;
                padding: 20px;
                border-radius: 10px;
                margin-bottom: 20px;
            }}
            .kpi-container {{
                display: flex;
                flex-wrap: wrap;
                gap: 15px;
                margin-bottom: 20px;
            }}
            .kpi-card {{
                background: #f5f5f5;
                border-radius: 10px;
                padding: 15px;
                flex: 1;
                min-width: 150px;
                text-align: center;
            }}
            .kpi-value {{
                font-size: 32px;
                font-weight: bold;
                color: #1E88E5;
            }}
            .kpi-label {{
                color: #666;
                font-size: 14px;
            }}
            .positive {{ color: #4CAF50; }}
            .negative {{ color: #F44336; }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin: 20px 0;
            }}
            th, td {{
                border: 1px solid #ddd;
                padding: 10px;
                text-align: left;
            }}
            th {{
                background: #1E88E5;
                color: white;
            }}
            tr:nth-child(even) {{
                background: #f9f9f9;
            }}
            .footer {{
                margin-top: 30px;
                padding: 15px;
                background: #f5f5f5;
                border-radius: 10px;
                font-size: 12px;
                color: #666;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>📊 大豐環保人力週報</h1>
            <p>報告日期：{today.strftime('%Y年%m月%d日')}</p>
        </div>
        
        <h2>📈 本週人力摘要</h2>
        <div class="kpi-container">
            <div class="kpi-card">
                <div class="kpi-value">{total_headcount}</div>
                <div class="kpi-label">目前在職人數</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value positive">+{weekly_changes['hires']}</div>
                <div class="kpi-label">本週入職</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value negative">-{weekly_changes['leaves']}</div>
                <div class="kpi-label">本週離職</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value">{monthly_turnover:.1f}%</div>
                <div class="kpi-label">本月離職率</div>
            </div>
        </div>
        
        {alerts_html}
        
        <h3>🏢 各部門人數統計</h3>
        <table>
            <thead>
                <tr>
                    <th>部門</th>
                    <th>總人數</th>
                    <th>直接</th>
                    <th>間接</th>
                </tr>
            </thead>
            <tbody>
                {dept_table_rows}
            </tbody>
        </table>
        
        <div class="footer">
            <p>此為系統自動產生的週報，如有問題請聯繫人資部門。</p>
            <p>大豐環保人力儀表板 | 產生時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
    </body>
    </html>
    """
    
    return html_content


def generate_plain_text_report(
    employees_df,
    departments_df,
    config: Dict[str, Any],
    alerts: List[Alert]
) -> str:
    """
    產生純文字格式週報
    """
    today = date.today()
    
    total_headcount = get_current_headcount(employees_df)
    weekly_changes = get_weekly_changes(employees_df)
    monthly_turnover = get_turnover_rate(employees_df, 'monthly')
    
    text_content = f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 大豐環保人力週報
報告日期：{today.strftime('%Y年%m月%d日')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📈 本週人力摘要
━━━━━━━━━━━━━━━━━━━━━━
  目前在職人數：{total_headcount} 人
  本週入職：+{weekly_changes['hires']} 人
  本週離職：-{weekly_changes['leaves']} 人
  淨變化：{weekly_changes['net_change']:+d} 人
  本月離職率：{monthly_turnover:.1f}%

"""
    
    if alerts:
        text_content += "⚠️ 異常警報\n━━━━━━━━━━━━━━━━━━━━━━\n"
        for alert in alerts:
            text_content += f"  • {alert.title}: {alert.message}\n"
        text_content += "\n"
    
    text_content += f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
此為系統自動產生的週報
大豐環保人力儀表板
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
    
    return text_content


def send_email(
    subject: str,
    html_body: str,
    text_body: str,
    config: Dict[str, Any],
    attachments: List[str] = None
) -> bool:
    """
    寄送 Email
    
    Args:
        subject: 郵件主旨
        html_body: HTML 內容
        text_body: 純文字內容
        config: 設定字典
        attachments: 附件檔案路徑列表
        
    Returns:
        是否成功
    """
    email_config = config.get('email', {})
    
    if not email_config.get('enabled', False):
        print("⚠️ Email 功能已停用")
        return False
    
    try:
        # 建立郵件
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = email_config.get('sender', '')
        msg['To'] = ', '.join(email_config.get('recipients', []))
        
        # 加入內容
        part1 = MIMEText(text_body, 'plain', 'utf-8')
        part2 = MIMEText(html_body, 'html', 'utf-8')
        msg.attach(part1)
        msg.attach(part2)
        
        # 加入附件
        if attachments:
            for file_path in attachments:
                if os.path.exists(file_path):
                    with open(file_path, 'rb') as f:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            'Content-Disposition',
                            f'attachment; filename={os.path.basename(file_path)}'
                        )
                        msg.attach(part)
        
        # 寄送
        smtp_server = email_config.get('smtp_server', '')
        smtp_port = email_config.get('smtp_port', 587)
        
        # 從環境變數取得密碼
        password = os.environ.get('EMAIL_PASSWORD', '')
        
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            if email_config.get('use_tls', True):
                server.starttls()
            
            if password:
                server.login(email_config.get('sender', ''), password)
            
            server.send_message(msg)
        
        print("✅ Email 已成功寄送")
        return True
        
    except Exception as e:
        print(f"❌ Email 寄送失敗：{e}")
        return False


def send_weekly_report(
    employees_df,
    departments_df,
    config: Dict[str, Any]
) -> bool:
    """
    寄送週報
    
    Args:
        employees_df: 員工資料 DataFrame
        departments_df: 部門對照表
        config: 設定字典
        
    Returns:
        是否成功
    """
    today = date.today()
    
    # 檢查警報
    alerts = check_all_alerts(employees_df, config, departments_df)
    
    # 產生報告內容
    html_content = generate_report_content(
        employees_df, departments_df, config, alerts
    )
    text_content = generate_plain_text_report(
        employees_df, departments_df, config, alerts
    )
    
    # 郵件主旨
    subject = f"[人力週報] {today.strftime('%Y年%m月%d日')} - 大豐環保"
    if alerts:
        subject = f"⚠️ {subject}"
    
    # 寄送
    return send_email(subject, html_content, text_content, config)
