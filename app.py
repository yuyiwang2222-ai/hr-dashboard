"""
大豐環保人力儀表板 - 主程式
============================================
Streamlit 互動式儀表板
執行方式：streamlit run app.py
"""

import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta
from calendar import monthrange
import os
import sys

# 設定路徑
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.data_loader import (
    load_config, load_employees, load_departments, 
    merge_employee_department, load_upcoming_departures, clear_cache, mask_name
)
from src.metrics import (
    get_current_headcount, get_weekly_changes, get_monthly_changes,
    get_turnover_rate, get_department_stats, get_business_unit_stats, get_position_stats,
    get_labor_type_stats, get_employment_type_stats, get_headcount_trend,
    get_headcount_at_month_end, get_headcount_at_date, get_semiannual_trend,
    get_period_changes
)
from src.charts import (
    create_trend_chart, create_department_chart, create_department_stacked_chart,
    create_position_pie, create_employment_type_pie, create_labor_type_pie
)
from src.alerts import check_all_alerts, Alert, AlertLevel, get_alerts_summary


# ============================================
# 頁面設定
# ============================================
st.set_page_config(
    page_title="大豐環保人力儀表板",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自訂 CSS（含手機響應式設計）
st.markdown("""
<style>
    /* 基本樣式 */
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E88E5;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
    }
    .alert-critical {
        background-color: #FFEBEE;
        border-left: 4px solid #F44336;
        padding: 1rem;
        margin: 0.5rem 0;
        border-radius: 4px;
    }
    .alert-warning {
        background-color: #FFF8E1;
        border-left: 4px solid #FFC107;
        padding: 1rem;
        margin: 0.5rem 0;
        border-radius: 4px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 2rem;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 1.1rem;
    }
    
    /* 手機響應式設計 */
    @media (max-width: 768px) {
        .main-header {
            font-size: 1.5rem;
        }
        .sub-header {
            font-size: 0.85rem;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 0.5rem;
            flex-wrap: wrap;
        }
        .stTabs [data-baseweb="tab"] {
            font-size: 0.9rem;
            padding: 0.5rem;
        }
        [data-testid="column"] {
            width: 50% !important;
            flex: 0 0 50% !important;
            min-width: 50% !important;
        }
        [data-testid="stMetric"] {
            padding: 0.5rem;
        }
        [data-testid="stMetric"] label {
            font-size: 0.75rem !important;
        }
        [data-testid="stMetric"] [data-testid="stMetricValue"] {
            font-size: 1.2rem !important;
        }
        .stRadio > div {
            flex-wrap: wrap;
        }
    }
    
    /* 更小螢幕 */
    @media (max-width: 480px) {
        .main-header {
            font-size: 1.2rem;
        }
        [data-testid="column"] {
            width: 100% !important;
            flex: 0 0 100% !important;
            min-width: 100% !important;
        }
    }
</style>
""", unsafe_allow_html=True)


# ============================================
# 密碼驗證
# ============================================
def check_password():
    """顯示登入表單並驗證密碼，回傳是否已通過驗證"""
    if st.session_state.get("authenticated"):
        return True

    # 登入頁面置中樣式
    st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none; }
    </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("## 🏢 大豐環保人力儀表板")
        st.markdown("請輸入密碼以存取儀表板")
        with st.form("login_form"):
            password = st.text_input("密碼", type="password", placeholder="請輸入密碼")
            submitted = st.form_submit_button("登入", use_container_width=True)
            if submitted:
                # 優先從 st.secrets 讀取，備案從 config.yaml 讀取
                try:
                    correct_password = st.secrets["auth"]["password"]
                except (KeyError, FileNotFoundError):
                    import yaml
                    with open("config.yaml", "r", encoding="utf-8") as f:
                        _cfg = yaml.safe_load(f)
                    correct_password = _cfg.get("auth", {}).get("password", "")
                if password == correct_password:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("❌ 密碼錯誤，請重新輸入")
    return False


if not check_password():
    st.stop()


# ============================================
# 載入資料
# ============================================
@st.cache_data(ttl=300)  # 5分鐘快取
def load_all_data():
    """載入所有資料"""
    config = load_config()
    employees = load_employees(config=config)  # 傳入 config 以支援外部資料載入
    departments = load_departments()
    
    if len(employees) == 0:
        return config, pd.DataFrame(), departments, [], True
    
    # 取得事業部列表（從資料中提取）
    if 'department' in employees.columns:
        business_units = sorted(employees[employees['status'] == '在職']['department'].dropna().unique().tolist())
    else:
        business_units = departments['name'].tolist()
    
    return config, employees, departments, business_units, False


config, employees_df, departments_df, business_units_list, data_empty = load_all_data()


# ============================================
# 側邊欄
# ============================================
with st.sidebar:
    st.image("https://via.placeholder.com/200x60/1E88E5/FFFFFF?text=大豐環保", width=200)
    st.markdown("---")
    
    # 重新載入按鈕
    if st.button("🔄 重新載入資料", use_container_width=True):
        clear_cache()  # 清除 data_loader 快取
        st.cache_data.clear()  # 清除 Streamlit 快取
        st.rerun()
    
    st.markdown("---")
    
    # 月份選擇（當月往前6個月）
    st.subheader("📅 查詢月份")
    
    # 生成過去6個月的選項
    today = date.today()
    month_options = []
    for i in range(6):  # 0 = 本月, 1 = 上個月, ..., 5 = 5個月前
        # 計算往前 i 個月的年月
        target_month = today.month - i
        target_year = today.year
        while target_month <= 0:
            target_month += 12
            target_year -= 1
        month_options.append((target_year, target_month, f"{target_year}年{target_month}月"))
    
    # 預設選擇本月
    selected_month_idx = st.selectbox(
        "選擇月份",
        range(len(month_options)),
        format_func=lambda i: month_options[i][2],
        key="month_selector"
    )
    selected_year, selected_month, _ = month_options[selected_month_idx]
    
    st.markdown("---")
    st.caption(f"資料更新時間：{datetime.now().strftime('%Y-%m-%d %H:%M')}")

    st.markdown("---")
    st.subheader("📧 週報寄送")
    is_local_windows = os.name == "nt"
    if not is_local_windows:
        st.info("目前為雲端環境，僅可查看預覽；寄送與附件檢查請在本機儀表板（http://127.0.0.1:8501）操作。")

    report_exists = False
    try:
        from send_report_email import get_email_preview, send_report_email

        preview = get_email_preview()
        recipients_text = "、".join(preview["recipients"])
        st.caption(f"收件人：{recipients_text}")
        st.caption(f"主旨：{preview['subject']}")
        if preview["report_exists"]:
            st.success(f"附件就緒：{preview['report_filename']}")
            report_exists = True
        else:
            st.warning(f"尚未找到附件：{preview['report_filename']}（請先產生 PDF）")
    except Exception as exc:
        st.error(f"❌ 無法載入寄信預覽：{exc}")

    email_confirmed = st.checkbox("我已確認內容，現在傳送週報", key="email_send_confirm")
    send_disabled = (not email_confirmed) or (not report_exists) or (not is_local_windows)
    if st.button("傳送", use_container_width=True, disabled=send_disabled):
        with st.spinner("正在傳送週報..."):
            try:
                sent = send_report_email()
                if sent:
                    st.success("✅ 週報已寄出")
                else:
                    st.error("❌ 週報寄送失敗，請檢查 Outlook、網路與報告檔案")
            except Exception as exc:
                st.error(f"❌ 週報寄送失敗：{exc}")


# ============================================
# 資料篩選
# ============================================
# 取得排除設定
system_conf = config.get('system', {})
excluded_depts = system_conf.get('excluded_departments', [])
excluded_keywords = system_conf.get('excluded_department_keywords', [])
exclude_blank = system_conf.get('exclude_blank_departments', True)

def filter_data(df, exclude_special_depts=True, exclude_part_time=True):
    """根據篩選條件過濾資料（用於在職人數計算）
    
    Args:
        df: 員工資料 DataFrame
        exclude_special_depts: 是否排除特殊部門（如再生紡織設計部、含「負責人」字眼、空白部門）
        exclude_part_time: 是否排除工讀生（預設排除）
    """
    filtered = df.copy()
    
    # 排除「排除」狀態（虛擬、雇主等）
    if 'status' in filtered.columns:
        filtered = filtered[filtered['status'] != '排除']
    
    # 排除特殊部門
    if exclude_special_depts:
        # 排除指定部門
        if excluded_depts:
            filtered = filtered[~filtered['department'].isin(excluded_depts)]
        
        # 排除部門名稱包含特定關鍵字（如「負責人」）
        if excluded_keywords:
            for keyword in excluded_keywords:
                # 檢查原始部門名稱
                if 'department_original' in filtered.columns:
                    filtered = filtered[~filtered['department_original'].astype(str).str.contains(keyword, na=False)]
                # 也檢查事業部名稱
                filtered = filtered[~filtered['department'].astype(str).str.contains(keyword, na=False)]
        
        # 排除空白/未分類部門
        if exclude_blank:
            filtered = filtered[filtered['department'].notna()]
            filtered = filtered[filtered['department'].astype(str).str.strip() != '']
            filtered = filtered[filtered['department'] != '未分類']
    
    # 排除工讀生（用於在職人數計算）
    if exclude_part_time and 'is_part_time' in filtered.columns:
        filtered = filtered[filtered['is_part_time'] == False]
    
    return filtered


def filter_department_data(df, department):
    """取得特定部門的資料（用於部門人員名單，包含工讀生）"""
    filtered = df.copy()
    
    # 排除「排除」狀態（虛擬、雇主等）
    if 'status' in filtered.columns:
        filtered = filtered[filtered['status'] != '排除']
    
    # 篩選部門
    filtered = filtered[filtered['department'] == department]
    
    return filtered


def style_hire_date(val):
    """到職日在未來時顯示紅色"""
    if pd.isna(val):
        return ''
    try:
        if isinstance(val, str):
            val = pd.to_datetime(val).date()
        if val > date.today():
            return 'color: red; font-weight: bold'
    except:
        pass
    return ''


# ============================================
# 主要內容
# ============================================

# 標題區
st.markdown('<p class="main-header">🏢 大豐環保人力儀表板</p>', unsafe_allow_html=True)
st.markdown(f'<p class="sub-header">即時人力變化追蹤 | 更新日期：{date.today().strftime("%Y年%m月%d日")}</p>', unsafe_allow_html=True)

# 檢查資料
if data_empty:
    st.error("⚠️ 找不到員工資料！請先執行 `python setup_data.py` 建立測試資料。")
    st.stop()

# 套用篩選
filtered_df = filter_data(employees_df)

# 頁籤
tab1, tab_staff, tab2, tab3, tab4 = st.tabs(["📊 首頁總覽", "👔 行政幕僚", "🏢 部門分析", "📋 異動名單", "⚠️ 警報看板"])

# ============================================
# Tab 1: 首頁總覽
# ============================================
with tab1:
    # 計算選擇月份的範圍
    query_month_start = date(selected_year, selected_month, 1)
    query_last_day = monthrange(selected_year, selected_month)[1]
    query_month_end = date(selected_year, selected_month, query_last_day)
    
    # 顯示查詢月份
    is_current_month = (selected_year == date.today().year and selected_month == date.today().month)
    month_label = f"{selected_year}年{selected_month}月" + ("（本月）" if is_current_month else "")
    
    # 計算該月月底人數（用於歷史月份）
    if is_current_month:
        # 本月使用目前在職人數
        total_headcount = get_current_headcount(filtered_df)
    else:
        # 歷史月份使用月底人數
        total_headcount = get_headcount_at_month_end(filtered_df, selected_year, selected_month)
    
    monthly_changes = get_monthly_changes(filtered_df, selected_year, selected_month)
    monthly_turnover = get_turnover_rate(filtered_df, 'monthly', selected_year, selected_month)
    
    # 計算與 2025年1月 的比較
    baseline_date = date(2025, 1, 1)  # 2025年1月1日
    baseline_headcount = get_headcount_at_date(filtered_df, baseline_date)
    headcount_diff_baseline = total_headcount - baseline_headcount
    
    # ===== 第1列：期間累計對比 =====
    st.markdown("##### 📈 期間累計（2025年1月至今）")
    
    # 取得即將離職人數
    upcoming_departures_all = load_upcoming_departures()
    upcoming_count_all = len(upcoming_departures_all)
    
    row1_col1, row1_col2, row1_col3 = st.columns(3)
    
    with row1_col1:
        delta_val = f"{monthly_changes['net_change']:+d}" if monthly_changes['net_change'] != 0 else None
        st.metric(
            label=f"👥 {month_label}在職人數",
            value=f"{total_headcount:,}",
            delta=delta_val
        )
    
    with row1_col2:
        # 與 2025年1月 比較
        st.metric(
            label="📊 與2025年1月比",
            value=f"{headcount_diff_baseline:+d}人",
            delta=f"基準：{baseline_headcount}人",
            delta_color="off"
        )
    
    with row1_col3:
        st.metric(
            label="📋 即將預計離職",
            value=f"{upcoming_count_all}人",
            delta=None
        )
    
    # ===== 第2列：當月數據（可展開/收合） =====
    with st.expander(f"📅 {selected_month}月（當月）數據", expanded=False):
        row2_col1, row2_col2, row2_col3, row2_col4 = st.columns(4)
        
        with row2_col1:
            st.metric(
                label=f"👥 {month_label}在職人數",
                value=f"{total_headcount:,}",
                delta=None
            )
        
        with row2_col2:
            st.metric(
                label=f"📥 {selected_month}月入職",
                value=f"+{monthly_changes['hires']}",
                delta=None
            )
        
        with row2_col3:
            st.metric(
                label=f"📤 {selected_month}月離職",
                value=f"-{monthly_changes['leaves']}",
                delta=None
            )
        
        with row2_col4:
            # 離職率顏色判斷
            turnover_delta = None
            if monthly_turnover > 10:
                turnover_delta = "危急"
            elif monthly_turnover > 5:
                turnover_delta = "偏高"
            
            st.metric(
                label=f"📉 {selected_month}月離職率",
                value=f"{monthly_turnover:.1f}%",
                delta=turnover_delta,
                delta_color="inverse" if turnover_delta else "off"
            )
    
    st.markdown("---")
    
    # 直條圖：集團及各部門人數比較（當月 vs 2025年1月）
    st.subheader("📊 人數比較（當月 vs 2025年1月）")
    
    # 計算各部門的基準和當月人數
    import plotly.graph_objects as go
    
    baseline_date = date(2025, 1, 1)
    comparison_data = []
    
    # 集團總計
    baseline_total = get_headcount_at_date(filtered_df, baseline_date)
    current_total = total_headcount
    comparison_data.append({
        'department': '集團總計',
        'baseline': baseline_total,
        'current': current_total,
        'diff': current_total - baseline_total
    })
    
    # 各部門
    departments = filtered_df['department'].unique()
    for dept in sorted(departments):
        # 排除不需要顯示的部門
        if dept == '再生處理部-宏偉':
            continue
        dept_data = filtered_df[filtered_df['department'] == dept]
        
        # 電子商務部排除移工
        if dept == '電子商務部' and 'is_migrant_worker' in dept_data.columns:
            dept_data = dept_data[dept_data['is_migrant_worker'] == False]
        dept_baseline = get_headcount_at_date(dept_data, baseline_date)
        
        # 計算當月人數
        dept_data_calc = dept_data.copy()
        dept_data_calc['hire_date'] = pd.to_datetime(dept_data_calc['hire_date']).dt.date
        dept_data_calc['leave_date'] = pd.to_datetime(dept_data_calc['leave_date']).dt.date
        
        if is_current_month:
            today = date.today()
            dept_current = len(dept_data_calc[
                (dept_data_calc['hire_date'] <= today) & 
                ((dept_data_calc['leave_date'].isna()) | (dept_data_calc['leave_date'] > today))
            ])
        else:
            dept_current = len(dept_data_calc[
                (dept_data_calc['hire_date'] <= query_month_end) & 
                ((dept_data_calc['leave_date'].isna()) | (dept_data_calc['leave_date'] > query_month_end))
            ])
        
        if dept_baseline > 0 or dept_current > 0:  # 只顯示有人數的部門
            comparison_data.append({
                'department': dept,
                'baseline': dept_baseline,
                'current': dept_current,
                'diff': dept_current - dept_baseline
            })
    
    # 按當月人數排序（集團總計永遠在最前面）
    comparison_data_sorted = [comparison_data[0]] + sorted(comparison_data[1:], key=lambda x: x['current'], reverse=True)
    
    # 創建直條圖
    fig = go.Figure()
    
    dept_names = [d['department'] for d in comparison_data_sorted]
    baseline_values = [d['baseline'] for d in comparison_data_sorted]
    current_values = [d['current'] for d in comparison_data_sorted]
    diff_values = [d['diff'] for d in comparison_data_sorted]
    
    # 2025年1月
    fig.add_trace(go.Bar(
        name='2025年1月',
        x=dept_names,
        y=baseline_values,
        marker_color='#9E9E9E',
        text=baseline_values,
        textposition='outside'
    ))
    
    # 當月
    fig.add_trace(go.Bar(
        name=f'{selected_year}年{selected_month}月',
        x=dept_names,
        y=current_values,
        marker_color='#1E88E5',
        text=current_values,
        textposition='outside'
    ))
    
    fig.update_layout(
        barmode='group',
        title=f'人數比較：2025年1月 vs {selected_year}年{selected_month}月',
        xaxis_title='部門',
        yaxis_title='人數',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        height=500
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # 顯示差異表格
    with st.expander("📋 查看詳細差異"):
        diff_df = pd.DataFrame(comparison_data_sorted)
        diff_df.columns = ['部門', '2025年1月', f'{selected_year}年{selected_month}月', '變化']
        diff_df['變化'] = diff_df['變化'].apply(lambda x: f"+{x}" if x > 0 else str(x))
        st.dataframe(diff_df, use_container_width=True, hide_index=True)
    
    # 直接人力比較表格（篩選：移工、外場人員、司機、其他、工讀生、作業員、無）
    with st.expander("📋 各事業部直接人力比較"):
        admin_data = []
        # 直接人力的職務類別
        direct_labor_categories = ['移工', '外場人員', '司機', '其他', '工讀生', '作業員', '無']
        
        # 取得直接人力資料（篩選指定職務類別）
        if 'job_category' in filtered_df.columns:
            admin_df = filtered_df[filtered_df['job_category'].isin(direct_labor_categories)]
        else:
            admin_df = filtered_df.copy()
        
        for dept in sorted(filtered_df['department'].unique()):
            if dept == '再生處理部-宏偉':
                continue
            dept_admin = admin_df[admin_df['department'] == dept]
            
            # 2025年1月非一線人員人數
            dept_admin_calc = dept_admin.copy()
            dept_admin_calc['hire_date'] = pd.to_datetime(dept_admin_calc['hire_date']).dt.date
            dept_admin_calc['leave_date'] = pd.to_datetime(dept_admin_calc['leave_date']).dt.date
            baseline_admin = len(dept_admin_calc[
                (dept_admin_calc['hire_date'] <= baseline_date) & 
                ((dept_admin_calc['leave_date'].isna()) | (dept_admin_calc['leave_date'] > baseline_date))
            ])
            
            # 當月非一線人員人數
            if is_current_month:
                today = date.today()
                current_admin = len(dept_admin_calc[
                    (dept_admin_calc['hire_date'] <= today) & 
                    ((dept_admin_calc['leave_date'].isna()) | (dept_admin_calc['leave_date'] > today))
                ])
            else:
                current_admin = len(dept_admin_calc[
                    (dept_admin_calc['hire_date'] <= query_month_end) & 
                    ((dept_admin_calc['leave_date'].isna()) | (dept_admin_calc['leave_date'] > query_month_end))
                ])
            
            if baseline_admin > 0 or current_admin > 0:
                # 計算變化百分比
                if baseline_admin > 0:
                    pct_change = ((current_admin - baseline_admin) / baseline_admin) * 100
                else:
                    pct_change = 100.0 if current_admin > 0 else 0.0
                
                admin_data.append({
                    'department': dept,
                    'baseline': baseline_admin,
                    'current': current_admin,
                    'diff': current_admin - baseline_admin,
                    'pct': pct_change
                })
        
        # 按當月人數排序
        admin_data_sorted = sorted(admin_data, key=lambda x: x['current'], reverse=True)
        
        admin_diff_df = pd.DataFrame(admin_data_sorted)
        admin_diff_df.columns = ['部門', '2025年1月', f'{selected_year}年{selected_month}月', '變化', '百分比']
        admin_diff_df['變化'] = admin_diff_df['變化'].apply(lambda x: f"+{x}" if x > 0 else str(x))
        admin_diff_df['百分比'] = admin_diff_df['百分比'].apply(lambda x: f"+{x:.1f}%" if x > 0 else f"{x:.1f}%")
        st.dataframe(admin_diff_df, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    
    # 下半部：圖表區
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.subheader("🏢 各部門人數")
        dept_stats = get_business_unit_stats(filtered_df)
        
        # 顯示模式選擇
        chart_mode = st.radio(
            "顯示模式",
            ["總人數", "直接/間接堆疊"],
            horizontal=True,
            key="dept_chart_mode"
        )
        
        if chart_mode == "直接/間接堆疊":
            dept_chart = create_department_stacked_chart(
                dept_stats, departments_df,
                config.get('display', {}).get('colors')
            )
        else:
            dept_chart = create_department_chart(
                dept_stats, departments_df, 'all',
                config.get('display', {}).get('colors')
            )
        st.plotly_chart(dept_chart, use_container_width=True)
    
    with col_right:
        st.subheader("📊 人員分布")
        
        # 分布類型選擇（移除員工狀態）
        dist_type = st.radio(
            "分布類型",
            ["職務分布", "直接/間接"],
            horizontal=True,
            key="dist_type"
        )
        
        if dist_type == "職務分布":
            position_stats = get_position_stats(filtered_df)
            # 只顯示前10大職務
            if len(position_stats) > 10:
                other_count = position_stats.iloc[10:]['count'].sum()
                position_stats = position_stats.head(10)
                position_stats = pd.concat([
                    position_stats,
                    pd.DataFrame([{'position': '其他', 'count': other_count}])
                ])
            pie_chart = create_position_pie(position_stats)
        else:
            labor_stats = get_labor_type_stats(filtered_df)
            pie_chart = create_labor_type_pie(
                labor_stats, config.get('display', {}).get('colors')
            )
        
        st.plotly_chart(pie_chart, use_container_width=True)
    
    # ===== 即將離職名單 =====
    with st.expander("📋 即將離職名單", expanded=False):
        upcoming_departures = load_upcoming_departures()
        if len(upcoming_departures) > 0:
            st.warning(f"共 {len(upcoming_departures)} 位同仁即將離職")
            st.dataframe(upcoming_departures, use_container_width=True, hide_index=True)
        else:
            st.info("目前沒有即將離職的人員")


# ============================================
# Tab 行政幕僚: 各事業部行政幕僚人員比較
# ============================================
with tab_staff:
    st.subheader("👔 各事業部行政幕僚人員比較(含前端會計)")
    st.caption("包含職務類別：行政幕僚、襄理、中階主管、副理級以上、業務")
    
    # 計算選擇月份的範圍
    query_month_start_staff = date(selected_year, selected_month, 1)
    query_last_day_staff = monthrange(selected_year, selected_month)[1]
    query_month_end_staff = date(selected_year, selected_month, query_last_day_staff)
    is_current_month_staff = (selected_year == date.today().year and selected_month == date.today().month)
    baseline_date_staff = date(2025, 1, 1)
    
    staff_data_tab = []
    # 行政幕僚相關職務類別
    staff_job_categories_tab = ['行政幕僚', '襄理', '中階主管', '副理級以上', '業務']
    
    # 取得行政幕僚人員資料
    if 'job_category' in filtered_df.columns:
        staff_df_tab = filtered_df[filtered_df['job_category'].isin(staff_job_categories_tab)]
    else:
        staff_df_tab = pd.DataFrame()
    
    for dept in sorted(filtered_df['department'].unique()):
        if dept == '再生處理部-宏偉':
            continue
        dept_staff_tab = staff_df_tab[staff_df_tab['department'] == dept]
        
        # 2025年1月行政幕僚人員人數
        dept_staff_calc_tab = dept_staff_tab.copy()
        dept_staff_calc_tab['hire_date'] = pd.to_datetime(dept_staff_calc_tab['hire_date']).dt.date
        dept_staff_calc_tab['leave_date'] = pd.to_datetime(dept_staff_calc_tab['leave_date']).dt.date
        baseline_staff_tab = len(dept_staff_calc_tab[
            (dept_staff_calc_tab['hire_date'] <= baseline_date_staff) & 
            ((dept_staff_calc_tab['leave_date'].isna()) | (dept_staff_calc_tab['leave_date'] > baseline_date_staff))
        ])
        
        # 人資部特例：2025年1月基準為11人（含已離職代主管宋承憲）
        if dept == '人資部':
            baseline_staff_tab = 11
        
        # 當月行政幕僚人員人數
        if is_current_month_staff:
            today_staff = date.today()
            current_staff_tab = len(dept_staff_calc_tab[
                (dept_staff_calc_tab['hire_date'] <= today_staff) & 
                ((dept_staff_calc_tab['leave_date'].isna()) | (dept_staff_calc_tab['leave_date'] > today_staff))
            ])
        else:
            current_staff_tab = len(dept_staff_calc_tab[
                (dept_staff_calc_tab['hire_date'] <= query_month_end_staff) & 
                ((dept_staff_calc_tab['leave_date'].isna()) | (dept_staff_calc_tab['leave_date'] > query_month_end_staff))
            ])
        
        if baseline_staff_tab > 0 or current_staff_tab > 0:
            # 計算變化百分比
            if baseline_staff_tab > 0:
                pct_change_tab = ((current_staff_tab - baseline_staff_tab) / baseline_staff_tab) * 100
            else:
                pct_change_tab = 100.0 if current_staff_tab > 0 else 0.0
            
            staff_data_tab.append({
                'department': dept,
                'baseline': baseline_staff_tab,
                'current': current_staff_tab,
                'diff': current_staff_tab - baseline_staff_tab,
                'pct': pct_change_tab
            })
    
    # 按當月人數排序
    staff_data_sorted_tab = sorted(staff_data_tab, key=lambda x: x['current'], reverse=True)
    
    # 計算總計（先計算，以便先顯示）
    total_baseline = sum([d['baseline'] for d in staff_data_tab])
    total_current = sum([d['current'] for d in staff_data_tab])
    total_diff = total_current - total_baseline
    total_pct = ((total_diff / total_baseline) * 100) if total_baseline > 0 else 0
    
    # 先顯示總計卡片（整體人數變化一目瞭然）
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("2025年1月總計", f"{total_baseline} 人")
    with col2:
        st.metric(f"{selected_year}年{selected_month}月總計", f"{total_current} 人")
    with col3:
        st.metric("總變化", f"{'+' if total_diff > 0 else ''}{total_diff} 人")
    with col4:
        st.metric("變化百分比", f"{'+' if total_pct > 0 else ''}{total_pct:.1f}%")
    
    st.markdown("---")
    
    # 再顯示各部門明細表格
    staff_diff_df_tab = pd.DataFrame(staff_data_sorted_tab)
    staff_diff_df_tab.columns = ['部門', '2025年1月', f'{selected_year}年{selected_month}月', '變化', '百分比']
    staff_diff_df_tab['變化'] = staff_diff_df_tab['變化'].apply(lambda x: f"+{x}" if x > 0 else str(x))
    staff_diff_df_tab['百分比'] = staff_diff_df_tab['百分比'].apply(lambda x: f"+{x:.1f}%" if x > 0 else f"{x:.1f}%")
    st.dataframe(staff_diff_df_tab, use_container_width=True, hide_index=True)


# ============================================
# Tab 2: 部門分析
# ============================================
with tab2:
    st.subheader("🏢 事業部詳細分析")
    
    # 事業部選擇
    dept_for_analysis = st.selectbox(
        "選擇事業部",
        business_units_list if business_units_list else ['無資料'],
        key="dept_analysis_select"
    )
    
    # 取得該部門所有員工（包含工讀生，用於名單顯示）
    dept_df_all = filter_department_data(employees_df, dept_for_analysis)
    # 取得該部門非工讀生（用於人數統計）
    dept_df = dept_df_all[dept_df_all['is_part_time'] == False] if 'is_part_time' in dept_df_all.columns else dept_df_all
    
    # 電子商務部額外排除移工
    if dept_for_analysis == '電子商務部' and 'is_migrant_worker' in dept_df.columns:
        dept_df = dept_df[dept_df['is_migrant_worker'] == False]
    
    # 部門 KPI（使用選擇的月份，不含工讀生）
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    if is_current_month:
        dept_headcount = get_current_headcount(dept_df)
    else:
        dept_headcount = get_headcount_at_month_end(dept_df, selected_year, selected_month)
    
    dept_monthly = get_monthly_changes(dept_df, selected_year, selected_month)
    dept_turnover = get_turnover_rate(dept_df, 'monthly', selected_year, selected_month)
    
    # 計算該部門與 2025年1月 的比較
    dept_baseline_date = date(2025, 1, 1)
    dept_baseline_headcount = get_headcount_at_date(dept_df, dept_baseline_date)
    dept_diff_baseline = dept_headcount - dept_baseline_headcount
    
    # 取得該部門即將離職人數
    upcoming_dept = load_upcoming_departures()
    dept_upcoming_count = len(upcoming_dept[upcoming_dept['事業部'] == dept_for_analysis]) if len(upcoming_dept) > 0 else 0
    
    with col1:
        st.metric("部門人數", dept_headcount)
    with col2:
        st.metric(f"{selected_month}月入職", f"+{dept_monthly['hires']}")
    with col3:
        st.metric(f"{selected_month}月離職", f"-{dept_monthly['leaves']}")
    with col4:
        st.metric(f"{selected_month}月離職率", f"{dept_turnover:.1f}%")
    with col5:
        st.metric("預計離職", f"{dept_upcoming_count}人")
    with col6:
        st.metric(
            label="與2025年1月比",
            value=f"{dept_diff_baseline:+d}人",
            delta=f"基準：{dept_baseline_headcount}人",
            delta_color="off"
        )
    
    # 備註說明
    st.markdown('''
    <p style="color: red; font-size: 0.9rem;">※ 工讀生不計算在總人數，僅羅列記錄</p>
    <p style="color: red; font-size: 0.9rem;">※ 2名身障名額列在總經理室</p>
    ''', unsafe_allow_html=True)
    
    st.markdown("---")
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.subheader("職務組成")
        dept_position_stats = get_position_stats(dept_df)
        if len(dept_position_stats) > 0:
            position_pie = create_position_pie(dept_position_stats)
            st.plotly_chart(position_pie, use_container_width=True)
        else:
            st.info("無資料")
    
    with col_right:
        st.subheader("直接/間接比例")
        dept_labor_stats = get_labor_type_stats(dept_df)
        if len(dept_labor_stats) > 0:
            labor_pie = create_labor_type_pie(
                dept_labor_stats, config.get('display', {}).get('colors')
            )
            st.plotly_chart(labor_pie, use_container_width=True)
        else:
            st.info("無資料")
    
    # 部門人員名單（顯示所有員工，包含工讀生）
    st.subheader("📋 部門人員名單")
    active_dept_employees = dept_df_all[dept_df_all['status'] == '在職'].copy()
    
    # 選擇要顯示的欄位
    display_cols = []
    col_names = {}
    if 'employee_id' in active_dept_employees.columns:
        display_cols.append('employee_id')
        col_names['employee_id'] = '員工編號'
    if 'name' in active_dept_employees.columns:
        display_cols.append('name')
        col_names['name'] = '姓名'
    if 'position' in active_dept_employees.columns:
        display_cols.append('position')
        col_names['position'] = '職務'
    if 'job_category' in active_dept_employees.columns:
        display_cols.append('job_category')
        col_names['job_category'] = '職務類別'
    if 'labor_type' in active_dept_employees.columns:
        display_cols.append('labor_type')
        col_names['labor_type'] = '直間接'
    if 'hire_date' in active_dept_employees.columns:
        display_cols.append('hire_date')
        col_names['hire_date'] = '到職日'
    
    if display_cols:
        display_df = active_dept_employees[display_cols].copy()
        # 姓名遮罩（個資保護）
        if 'name' in display_cols:
            display_df['name'] = display_df['name'].apply(mask_name)
        display_df.columns = [col_names.get(c, c) for c in display_cols]
        
        # 套用到職日紅字樣式（未來日期）
        if '到職日' in display_df.columns:
            styled_df = display_df.style.map(
                style_hire_date, subset=['到職日']
            )
            st.dataframe(styled_df, use_container_width=True, hide_index=True)
        else:
            st.dataframe(display_df, use_container_width=True, hide_index=True)
    else:
        st.info("無資料")
    
    # ===== 該部門即將離職名單 =====
    with st.expander(f"📋 {dept_for_analysis} 即將離職名單", expanded=False):
        upcoming_departures = load_upcoming_departures()
        # 篩選該事業部
        dept_upcoming = upcoming_departures[upcoming_departures['事業部'] == dept_for_analysis].copy() if len(upcoming_departures) > 0 else upcoming_departures
        if len(dept_upcoming) > 0:
            # 姓名遮罩（個資保護）
            if '姓名' in dept_upcoming.columns:
                dept_upcoming['姓名'] = dept_upcoming['姓名'].apply(mask_name)
            st.warning(f"共 {len(dept_upcoming)} 位同仁即將離職")
            st.dataframe(dept_upcoming, use_container_width=True, hide_index=True)
        else:
            st.info(f"{dept_for_analysis} 目前沒有即將離職的人員")


# ============================================
# Tab 3: 異動名單
# ============================================
with tab3:
    st.subheader("📋 人員異動名單")
    
    # ===== 即將離職名單（全公司） =====
    with st.expander("📋 即將離職名單", expanded=True):
        upcoming_departures = load_upcoming_departures()
        if len(upcoming_departures) > 0:
            # 姓名遮罩（個資保護）
            upcoming_display = upcoming_departures.copy()
            if '姓名' in upcoming_display.columns:
                upcoming_display['姓名'] = upcoming_display['姓名'].apply(mask_name)
            st.warning(f"共 {len(upcoming_display)} 位同仁即將離職")
            st.dataframe(upcoming_display, use_container_width=True, hide_index=True)
        else:
            st.info("目前沒有即將離職的人員")
    
    st.markdown("---")
    
    # 期間選擇
    period_option = st.radio(
        "選擇期間",
        ["本月", "自訂"],
        horizontal=True
    )
    
    if period_option == "本月":
        start_date = date.today().replace(day=1)
        end_date = date.today()
    else:
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("開始日期", date.today() - timedelta(days=30))
        with col2:
            end_date = st.date_input("結束日期", date.today())
    
    st.markdown(f"**期間：{start_date} ~ {end_date}**")
    
    # 計算異動
    from src.metrics import get_period_changes
    changes = get_period_changes(filtered_df, start_date, end_date)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader(f"📥 入職名單（{changes['hires']} 人）")
        if changes['hires'] > 0:
            hire_list = changes['hire_list'].copy()
            # 選擇要顯示的欄位
            display_cols = ['employee_id', 'name', 'department', 'position', 'hire_date']
            display_cols = [c for c in display_cols if c in hire_list.columns]
            hire_list = hire_list[display_cols]
            # 姓名遮罩（個資保護）
            if 'name' in display_cols:
                name_idx = display_cols.index('name')
                hire_list.iloc[:, name_idx] = hire_list.iloc[:, name_idx].apply(mask_name)
            hire_list.columns = ['員工編號', '姓名', '事業部', '職務', '到職日'][:len(display_cols)]
            
            # 套用到職日紅字樣式（未來日期）
            if '到職日' in hire_list.columns:
                styled_hire = hire_list.style.map(style_hire_date, subset=['到職日'])
                st.dataframe(styled_hire, use_container_width=True, hide_index=True)
            else:
                st.dataframe(hire_list, use_container_width=True, hide_index=True)
        else:
            st.info("此期間無入職人員")
    
    with col2:
        st.subheader(f"📤 離職名單（{changes['leaves']} 人）")
        if changes['leaves'] > 0:
            leave_list = changes['leave_list'].copy()
            # 選擇要顯示的欄位
            display_cols = ['employee_id', 'name', 'department', 'position', 'leave_date']
            display_cols = [c for c in display_cols if c in leave_list.columns]
            leave_list = leave_list[display_cols]
            # 姓名遮罩（個資保護）
            if 'name' in display_cols:
                name_idx = display_cols.index('name')
                leave_list.iloc[:, name_idx] = leave_list.iloc[:, name_idx].apply(mask_name)
            leave_list.columns = ['員工編號', '姓名', '事業部', '職務', '離職日'][:len(display_cols)]
            st.dataframe(leave_list, use_container_width=True, hide_index=True)
        else:
            st.info("此期間無離職人員")


# ============================================
# Tab 4: 警報看板
# ============================================
with tab4:
    st.subheader("⚠️ 異常警報看板")
    
    # 檢查警報
    alerts = check_all_alerts(employees_df, config, departments_df)
    
    # 摘要
    if alerts:
        critical_count = sum(1 for a in alerts if a.level == AlertLevel.CRITICAL)
        warning_count = sum(1 for a in alerts if a.level == AlertLevel.WARNING)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("總警報數", len(alerts))
        with col2:
            st.metric("🔴 危急", critical_count)
        with col3:
            st.metric("🟡 警告", warning_count)
        
        st.markdown("---")
        
        # 警報列表
        for alert in alerts:
            if alert.level == AlertLevel.CRITICAL:
                st.markdown(f"""
                <div class="alert-critical">
                    <strong>{alert.title}</strong><br>
                    {alert.message}<br>
                    <small>當前值：{alert.current_value} | 閾值：{alert.threshold} | 時間：{alert.timestamp.strftime('%H:%M')}</small>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class="alert-warning">
                    <strong>{alert.title}</strong><br>
                    {alert.message}<br>
                    <small>當前值：{alert.current_value} | 閾值：{alert.threshold} | 時間：{alert.timestamp.strftime('%H:%M')}</small>
                </div>
                """, unsafe_allow_html=True)
    else:
        st.success("✅ 目前無異常警報，一切正常！")
    
    st.markdown("---")
    
    # 警報設定說明
    with st.expander("📋 警報閾值設定"):
        alert_config = config.get('alerts', {})
        st.write("**月離職率**")
        st.write(f"- 黃燈：> {alert_config.get('monthly_turnover_rate', {}).get('warning', 5)}%")
        st.write(f"- 紅燈：> {alert_config.get('monthly_turnover_rate', {}).get('critical', 10)}%")
        st.write("**部門月離職人數**")
        st.write(f"- 黃燈：> {alert_config.get('department_monthly_leaves', {}).get('warning', 3)} 人")
        st.write("**週人數變化率**")
        st.write(f"- 黃燈：> {alert_config.get('weekly_headcount_change', {}).get('warning', 5)}%")


# ============================================
# 頁尾
# ============================================
st.markdown("---")
st.caption("大豐環保人力儀表板 © 2026 | 如有問題請聯繫資訊部")
