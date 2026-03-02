"""
大豐環保人力儀表板 - PDF 報告生成
============================================
生成人力變化分析 PDF 報告
執行方式：python generate_pdf_report.py
"""

import os
import sys
from datetime import date, datetime
from calendar import monthrange

# 設定路徑
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import pandas as pd
import matplotlib
matplotlib.use('Agg')  # 非互動式後端
import matplotlib.pyplot as plt

# 註冊中文字體
try:
    # Windows 微軟正黑體
    pdfmetrics.registerFont(TTFont('MicrosoftJhengHei', 'C:/Windows/Fonts/msjh.ttc'))
    FONT_NAME = 'MicrosoftJhengHei'
except:
    try:
        # 備用：標楷體
        pdfmetrics.registerFont(TTFont('DFKai', 'C:/Windows/Fonts/kaiu.ttf'))
        FONT_NAME = 'DFKai'
    except:
        FONT_NAME = 'Helvetica'

# Matplotlib 中文字體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'DFKai-SB', 'SimHei']
plt.rcParams['axes.unicode_minus'] = False

from src.data_loader import (
    load_config, load_employees, load_departments, 
    load_upcoming_departures, clear_cache
)
from src.metrics import (
    get_current_headcount, get_monthly_changes,
    get_turnover_rate, get_department_stats, 
    get_headcount_at_date, get_headcount_at_month_end
)


def filter_data(df, config):
    """根據篩選條件過濾資料"""
    filtered = df.copy()
    
    system_conf = config.get('system', {})
    excluded_depts = system_conf.get('excluded_departments', [])
    excluded_keywords = system_conf.get('excluded_department_keywords', [])
    
    # 排除「排除」狀態
    if 'status' in filtered.columns:
        filtered = filtered[filtered['status'] != '排除']
    
    # 排除指定部門
    if excluded_depts:
        filtered = filtered[~filtered['department'].isin(excluded_depts)]
    
    # 排除關鍵字部門
    if excluded_keywords:
        for keyword in excluded_keywords:
            if 'department_original' in filtered.columns:
                filtered = filtered[~filtered['department_original'].astype(str).str.contains(keyword, na=False)]
            filtered = filtered[~filtered['department'].astype(str).str.contains(keyword, na=False)]
    
    # 排除空白部門
    filtered = filtered[filtered['department'].notna()]
    filtered = filtered[filtered['department'].astype(str).str.strip() != '']
    filtered = filtered[filtered['department'] != '未分類']
    
    # 排除工讀生
    if 'is_part_time' in filtered.columns:
        filtered = filtered[filtered['is_part_time'] == False]
    
    return filtered


def create_summary_table(filtered_df, config, selected_year, selected_month):
    """建立摘要表格數據"""
    baseline_date = date(2025, 1, 1)
    
    # 判斷是否為當月
    today = date.today()
    is_current_month = (selected_year == today.year and selected_month == today.month)
    
    # 計算當月最後一天
    last_day = monthrange(selected_year, selected_month)[1]
    query_month_end = date(selected_year, selected_month, last_day)
    
    # 計算總人數
    if is_current_month:
        total_headcount = get_current_headcount(filtered_df)
    else:
        total_headcount = get_headcount_at_month_end(filtered_df, selected_year, selected_month)
    
    # 計算月度變化
    monthly_changes = get_monthly_changes(filtered_df, selected_year, selected_month)
    monthly_turnover = get_turnover_rate(filtered_df, 'monthly', selected_year, selected_month)
    
    # 與基準比較
    baseline_headcount = get_headcount_at_date(filtered_df, baseline_date)
    headcount_diff = total_headcount - baseline_headcount
    
    return {
        'total': total_headcount,
        'hires': monthly_changes['hires'],
        'leaves': monthly_changes['leaves'],
        'turnover': monthly_turnover,
        'baseline': baseline_headcount,
        'diff': headcount_diff
    }


def create_department_comparison(filtered_df, selected_year, selected_month):
    """建立各部門比較數據"""
    baseline_date = date(2025, 1, 1)
    today = date.today()
    is_current_month = (selected_year == today.year and selected_month == today.month)
    last_day = monthrange(selected_year, selected_month)[1]
    query_month_end = date(selected_year, selected_month, last_day)
    
    comparison_data = []
    
    for dept in sorted(filtered_df['department'].unique()):
        if dept == '再生處理部-宏偉':
            continue
        
        dept_data = filtered_df[filtered_df['department'] == dept].copy()
        
        # 電子商務部排除移工
        if dept == '電子商務部' and 'is_migrant_worker' in dept_data.columns:
            dept_data = dept_data[dept_data['is_migrant_worker'] == False]
        
        dept_baseline = get_headcount_at_date(dept_data, baseline_date)
        
        # 計算當月人數
        dept_data['hire_date'] = pd.to_datetime(dept_data['hire_date']).dt.date
        dept_data['leave_date'] = pd.to_datetime(dept_data['leave_date']).dt.date
        
        if is_current_month:
            dept_current = len(dept_data[
                (dept_data['hire_date'] <= today) & 
                ((dept_data['leave_date'].isna()) | (dept_data['leave_date'] > today))
            ])
        else:
            dept_current = len(dept_data[
                (dept_data['hire_date'] <= query_month_end) & 
                ((dept_data['leave_date'].isna()) | (dept_data['leave_date'] > query_month_end))
            ])
        
        if dept_baseline > 0 or dept_current > 0:
            comparison_data.append({
                'department': dept,
                'baseline': dept_baseline,
                'current': dept_current,
                'diff': dept_current - dept_baseline
            })
    
    return sorted(comparison_data, key=lambda x: x['current'], reverse=True)


def create_staff_comparison(filtered_df, selected_year, selected_month):
    """建立行政幕僚人員比較數據"""
    baseline_date = date(2025, 1, 1)
    today = date.today()
    is_current_month = (selected_year == today.year and selected_month == today.month)
    last_day = monthrange(selected_year, selected_month)[1]
    query_month_end = date(selected_year, selected_month, last_day)
    
    staff_job_categories = ['行政幕僚', '襄理', '中階主管', '副理級以上', '業務']
    
    if 'job_category' in filtered_df.columns:
        staff_df = filtered_df[filtered_df['job_category'].isin(staff_job_categories)]
    else:
        return []
    
    staff_data = []
    
    for dept in sorted(filtered_df['department'].unique()):
        if dept == '再生處理部-宏偉':
            continue
        
        dept_staff = staff_df[staff_df['department'] == dept].copy()
        dept_staff['hire_date'] = pd.to_datetime(dept_staff['hire_date']).dt.date
        dept_staff['leave_date'] = pd.to_datetime(dept_staff['leave_date']).dt.date
        
        baseline_staff = len(dept_staff[
            (dept_staff['hire_date'] <= baseline_date) & 
            ((dept_staff['leave_date'].isna()) | (dept_staff['leave_date'] > baseline_date))
        ])
        
        # 人資部特例
        if dept == '人資部':
            baseline_staff = 11
        
        if is_current_month:
            current_staff = len(dept_staff[
                (dept_staff['hire_date'] <= today) & 
                ((dept_staff['leave_date'].isna()) | (dept_staff['leave_date'] > today))
            ])
        else:
            current_staff = len(dept_staff[
                (dept_staff['hire_date'] <= query_month_end) & 
                ((dept_staff['leave_date'].isna()) | (dept_staff['leave_date'] > query_month_end))
            ])
        
        if baseline_staff > 0 or current_staff > 0:
            pct = ((current_staff - baseline_staff) / baseline_staff * 100) if baseline_staff > 0 else (100.0 if current_staff > 0 else 0.0)
            staff_data.append({
                'department': dept,
                'baseline': baseline_staff,
                'current': current_staff,
                'diff': current_staff - baseline_staff,
                'pct': pct
            })
    
    return sorted(staff_data, key=lambda x: x['current'], reverse=True)


def create_non_frontline_comparison(filtered_df, selected_year, selected_month):
    """建立非行政幕僚比較數據"""
    baseline_date = date(2025, 1, 1)
    today = date.today()
    is_current_month = (selected_year == today.year and selected_month == today.month)
    last_day = monthrange(selected_year, selected_month)[1]
    query_month_end = date(selected_year, selected_month, last_day)
    
    non_admin_categories = ['移工', '外場人員', '司機', '其他', '工讀生', '作業員', '無']
    
    if 'job_category' in filtered_df.columns:
        admin_df = filtered_df[filtered_df['job_category'].isin(non_admin_categories)]
    else:
        admin_df = filtered_df.copy()
    
    admin_data = []
    
    for dept in sorted(filtered_df['department'].unique()):
        if dept == '再生處理部-宏偉':
            continue
        
        dept_admin = admin_df[admin_df['department'] == dept].copy()
        dept_admin['hire_date'] = pd.to_datetime(dept_admin['hire_date']).dt.date
        dept_admin['leave_date'] = pd.to_datetime(dept_admin['leave_date']).dt.date
        
        baseline_admin = len(dept_admin[
            (dept_admin['hire_date'] <= baseline_date) & 
            ((dept_admin['leave_date'].isna()) | (dept_admin['leave_date'] > baseline_date))
        ])
        
        if is_current_month:
            current_admin = len(dept_admin[
                (dept_admin['hire_date'] <= today) & 
                ((dept_admin['leave_date'].isna()) | (dept_admin['leave_date'] > today))
            ])
        else:
            current_admin = len(dept_admin[
                (dept_admin['hire_date'] <= query_month_end) & 
                ((dept_admin['leave_date'].isna()) | (dept_admin['leave_date'] > query_month_end))
            ])
        
        if baseline_admin > 0 or current_admin > 0:
            pct = ((current_admin - baseline_admin) / baseline_admin * 100) if baseline_admin > 0 else (100.0 if current_admin > 0 else 0.0)
            admin_data.append({
                'department': dept,
                'baseline': baseline_admin,
                'current': current_admin,
                'diff': current_admin - baseline_admin,
                'pct': pct
            })
    
    return sorted(admin_data, key=lambda x: x['current'], reverse=True)


def create_bar_chart(comparison_data, selected_year, selected_month, output_path):
    """建立部門比較長條圖"""
    fig, ax = plt.subplots(figsize=(12, 6))
    
    depts = [d['department'] for d in comparison_data[:12]]  # 最多顯示12個
    baseline = [d['baseline'] for d in comparison_data[:12]]
    current = [d['current'] for d in comparison_data[:12]]
    
    x = range(len(depts))
    width = 0.35
    
    bars1 = ax.bar([i - width/2 for i in x], baseline, width, label='2025年1月', color='#9E9E9E')
    bars2 = ax.bar([i + width/2 for i in x], current, width, label=f'{selected_year}年{selected_month}月', color='#1E88E5')
    
    ax.set_ylabel('人數')
    ax.set_title(f'各部門人數比較：2025年1月 vs {selected_year}年{selected_month}月')
    ax.set_xticks(x)
    ax.set_xticklabels(depts, rotation=45, ha='right')
    ax.legend()
    
    # 數值標籤
    for bar in bars1:
        height = bar.get_height()
        ax.annotate(f'{int(height)}', xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=8)
    for bar in bars2:
        height = bar.get_height()
        ax.annotate(f'{int(height)}', xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=8)
    
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    
    return output_path


def generate_pdf_report(output_path=None):
    """生成 PDF 報告"""
    
    # 載入資料
    print("📊 載入資料中...")
    config = load_config()
    employees_df = load_employees(config=config)
    
    if len(employees_df) == 0:
        print("❌ 無法載入員工資料")
        return None
    
    filtered_df = filter_data(employees_df, config)
    
    # 取得當前日期
    today = date.today()
    selected_year = today.year
    selected_month = today.month
    
    # 設定輸出路徑（存到「報告」資料夾）
    if output_path is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        report_dir = os.path.join(script_dir, "報告")
        os.makedirs(report_dir, exist_ok=True)
        output_path = os.path.join(report_dir, f"人力分析報告_{today.strftime('%Y%m%d')}.pdf")
    
    print(f"📝 生成報告: {output_path}")
    
    # 建立 PDF
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=1.5*cm,
        leftMargin=1.5*cm,
        topMargin=1.5*cm,
        bottomMargin=1.5*cm
    )
    
    # 樣式
    styles = getSampleStyleSheet()
    
    # 自訂樣式
    title_style = ParagraphStyle(
        'ChineseTitle',
        parent=styles['Title'],
        fontName=FONT_NAME,
        fontSize=20,
        spaceAfter=20,
        alignment=1  # 置中
    )
    
    heading_style = ParagraphStyle(
        'ChineseHeading',
        parent=styles['Heading2'],
        fontName=FONT_NAME,
        fontSize=14,
        spaceBefore=15,
        spaceAfter=10
    )
    
    normal_style = ParagraphStyle(
        'ChineseNormal',
        parent=styles['Normal'],
        fontName=FONT_NAME,
        fontSize=10,
        spaceAfter=6
    )
    
    # 內容
    story = []
    
    # 標題
    story.append(Paragraph("大豐環保人力分析報告", title_style))
    story.append(Paragraph(f"報告日期：{today.strftime('%Y年%m月%d日')}", normal_style))
    story.append(Spacer(1, 20))
    
    # ===== 1. 摘要 =====
    story.append(Paragraph("一、人力概況摘要", heading_style))
    
    summary = create_summary_table(filtered_df, config, selected_year, selected_month)
    
    summary_data = [
        ['指標', '數值'],
        [f'{selected_month}月在職人數', f'{summary["total"]} 人'],
        [f'{selected_month}月入職', f'+{summary["hires"]} 人'],
        [f'{selected_month}月離職', f'-{summary["leaves"]} 人'],
        [f'{selected_month}月離職率', f'{summary["turnover"]:.1f}%'],
        ['2025年1月基準', f'{summary["baseline"]} 人'],
        ['與基準比較', f'{summary["diff"]:+d} 人'],
    ]
    
    summary_table = Table(summary_data, colWidths=[6*cm, 4*cm])
    summary_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E88E5')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5F5')]),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 20))
    
    # ===== 2. 各部門人數比較 =====
    story.append(Paragraph("二、各部門人數比較", heading_style))
    
    comparison_data = create_department_comparison(filtered_df, selected_year, selected_month)
    
    # 建立圖表
    chart_path = f"temp_chart_{today.strftime('%Y%m%d')}.png"
    create_bar_chart(comparison_data, selected_year, selected_month, chart_path)
    
    if os.path.exists(chart_path):
        story.append(Image(chart_path, width=16*cm, height=8*cm))
        story.append(Spacer(1, 10))
    
    # 表格
    dept_table_data = [['部門', '2025年1月', f'{selected_year}年{selected_month}月', '變化']]
    for d in comparison_data:
        diff_str = f"+{d['diff']}" if d['diff'] > 0 else str(d['diff'])
        dept_table_data.append([d['department'], d['baseline'], d['current'], diff_str])
    
    dept_table = Table(dept_table_data, colWidths=[5*cm, 3*cm, 3*cm, 3*cm])
    dept_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E88E5')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F5F5F5')]),
    ]))
    story.append(dept_table)
    
    story.append(PageBreak())
    
    # ===== 3. 行政幕僚人員比較 =====
    story.append(Paragraph("三、各事業部行政幕僚人員比較(含前端會計)", heading_style))
    story.append(Paragraph("包含職務類別：行政幕僚、襄理、中階主管、副理級以上、業務", normal_style))
    
    staff_data = create_staff_comparison(filtered_df, selected_year, selected_month)
    
    staff_table_data = [['部門', '2025年1月', f'{selected_year}年{selected_month}月', '變化', '百分比']]
    for d in staff_data:
        diff_str = f"+{d['diff']}" if d['diff'] > 0 else str(d['diff'])
        pct_str = f"+{d['pct']:.1f}%" if d['pct'] > 0 else f"{d['pct']:.1f}%"
        staff_table_data.append([d['department'], d['baseline'], d['current'], diff_str, pct_str])
    
    staff_table = Table(staff_table_data, colWidths=[4*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm])
    staff_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4CAF50')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#E8F5E9')]),
    ]))
    story.append(staff_table)
    story.append(Spacer(1, 20))
    
    # ===== 4. 直接人力比較 =====
    story.append(Paragraph("四、各事業部直接人力比較", heading_style))
    story.append(Paragraph("篩選職務類別：移工、外場人員、司機、工讀生、作業員", normal_style))
    
    non_frontline_data = create_non_frontline_comparison(filtered_df, selected_year, selected_month)
    
    nf_table_data = [['部門', '2025年1月', f'{selected_year}年{selected_month}月', '變化', '百分比']]
    for d in non_frontline_data:
        diff_str = f"+{d['diff']}" if d['diff'] > 0 else str(d['diff'])
        pct_str = f"+{d['pct']:.1f}%" if d['pct'] > 0 else f"{d['pct']:.1f}%"
        nf_table_data.append([d['department'], d['baseline'], d['current'], diff_str, pct_str])
    
    nf_table = Table(nf_table_data, colWidths=[4*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm])
    nf_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FF9800')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#FFF8E1')]),
    ]))
    story.append(nf_table)
    story.append(Spacer(1, 20))
    
    # ===== 5. 即將離職名單 =====
    story.append(Paragraph("五、預計離職名單", heading_style))
    
    upcoming = load_upcoming_departures()
    if len(upcoming) > 0:
        story.append(Paragraph(f"共 {len(upcoming)} 位同仁預計離職", normal_style))
        
        upcoming_table_data = [['事業部', '姓名', '職稱', '離職日']]
        for _, row in upcoming.iterrows():
            leave_date = row['離職日'].strftime('%Y/%m/%d') if hasattr(row['離職日'], 'strftime') else str(row['離職日'])
            upcoming_table_data.append([row['事業部'], row['姓名'], row['職稱'], leave_date])
        
        upcoming_table = Table(upcoming_table_data, colWidths=[4*cm, 3*cm, 4*cm, 3*cm])
        upcoming_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F44336')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#FFEBEE')]),
        ]))
        story.append(upcoming_table)
    else:
        story.append(Paragraph("目前沒有預計離職的人員", normal_style))
    
    # 建立 PDF
    doc.build(story)
    
    # 清理暫存圖檔
    if os.path.exists(chart_path):
        os.remove(chart_path)
    
    print(f"✅ 報告已生成: {output_path}")
    return output_path


if __name__ == "__main__":
    # 切換到專案目錄
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    # 生成報告
    output = generate_pdf_report()
    
    if output:
        # 自動開啟 PDF
        os.startfile(output)
