"""
大豐環保人力儀表板 - 指標計算模組
============================================
負責計算各項人力指標
"""

import pandas as pd
from datetime import date, datetime, timedelta
from calendar import monthrange
from typing import Dict, Optional, Any, List


def get_active_employees(df: pd.DataFrame, as_of_date: date = None) -> pd.DataFrame:
    """取得在職員工（排除到職日在未來的人員）
    
    Args:
        df: 員工資料 DataFrame
        as_of_date: 基準日期（預設為今日）
        
    Returns:
        在職員工 DataFrame
    """
    if as_of_date is None:
        as_of_date = date.today()
    
    active = df[df['status'] == '在職'].copy()
    
    # 確保 hire_date 是日期格式
    if 'hire_date' in active.columns:
        active['hire_date'] = pd.to_datetime(active['hire_date']).dt.date
        # 排除到職日在未來的人員
        active = active[active['hire_date'] <= as_of_date]
    
    return active


def get_current_headcount(
    df: pd.DataFrame, 
    filters: Optional[Dict[str, Any]] = None
) -> int:
    """
    計算目前在職人數
    
    Args:
        df: 員工資料 DataFrame
        filters: 篩選條件字典，例如 {'department': 'IT', 'labor_type': '直接'}
        
    Returns:
        在職人數
    """
    active = get_active_employees(df)
    
    if filters:
        for key, value in filters.items():
            if key in active.columns and value is not None:
                if isinstance(value, list):
                    active = active[active[key].isin(value)]
                else:
                    active = active[active[key] == value]
    
    return len(active)


def get_period_changes(
    df: pd.DataFrame, 
    start_date: date, 
    end_date: date
) -> Dict[str, Any]:
    """
    計算期間內的人員異動
    
    Args:
        df: 員工資料 DataFrame
        start_date: 開始日期
        end_date: 結束日期
        
    Returns:
        異動統計字典：{'hires': int, 'leaves': int, 'net_change': int, 
                       'hire_list': DataFrame, 'leave_list': DataFrame}
    """
    # 轉換日期欄位
    df = df.copy()
    df['hire_date'] = pd.to_datetime(df['hire_date']).dt.date
    df['leave_date'] = pd.to_datetime(df['leave_date']).dt.date
    
    # 入職名單
    hire_mask = (df['hire_date'] >= start_date) & (df['hire_date'] <= end_date)
    hire_list = df[hire_mask]
    
    # 離職名單
    leave_mask = (df['leave_date'] >= start_date) & (df['leave_date'] <= end_date)
    leave_list = df[leave_mask]
    
    hires = len(hire_list)
    leaves = len(leave_list)
    
    return {
        'hires': hires,
        'leaves': leaves,
        'net_change': hires - leaves,
        'hire_list': hire_list,
        'leave_list': leave_list
    }


def get_weekly_changes(df: pd.DataFrame) -> Dict[str, Any]:
    """取得本週異動"""
    today = date.today()
    # 計算本週一
    start_of_week = today - timedelta(days=today.weekday())
    return get_period_changes(df, start_of_week, today)


def get_monthly_changes(df: pd.DataFrame, year: int = None, month: int = None) -> Dict[str, Any]:
    """取得指定月份異動（月基準為該月最後一日）
    
    Args:
        df: 員工資料 DataFrame
        year: 年份（預設為今年）
        month: 月份（預設為本月）
    """
    today = date.today()
    
    if year is None:
        year = today.year
    if month is None:
        month = today.month
    
    start_of_month = date(year, month, 1)
    # 取得該月最後一日
    last_day = monthrange(year, month)[1]
    end_of_month = date(year, month, last_day)
    
    # 如果選擇的是當前月份且還沒到月底，使用今日為截止日
    if year == today.year and month == today.month:
        end_date = min(today, end_of_month)
    else:
        end_date = end_of_month
    
    return get_period_changes(df, start_of_month, end_date)


def get_headcount_at_date(df: pd.DataFrame, target_date: date) -> int:
    """
    計算指定日期的在職人數
    
    Args:
        df: 員工資料 DataFrame
        target_date: 目標日期
        
    Returns:
        該日期的在職人數
    """
    df = df.copy()
    df['hire_date'] = pd.to_datetime(df['hire_date']).dt.date
    df['leave_date'] = pd.to_datetime(df['leave_date']).dt.date
    
    # 在職條件：hire_date <= target_date AND (leave_date is null OR leave_date > target_date)
    active_count = len(df[
        (df['hire_date'] <= target_date) & 
        ((df['leave_date'].isna()) | (df['leave_date'] > target_date))
    ])
    
    return active_count


def get_headcount_at_month_end(df: pd.DataFrame, year: int, month: int) -> int:
    """
    計算指定月份月底的在職人數
    
    Args:
        df: 員工資料 DataFrame
        year: 年份
        month: 月份
        
    Returns:
        該月底的在職人數
    """
    today = date.today()
    last_day = monthrange(year, month)[1]
    month_end = date(year, month, last_day)
    
    # 如果是未來日期，使用今日
    if month_end > today:
        month_end = today
    
    return get_headcount_at_date(df, month_end)


def get_turnover_rate(
    df: pd.DataFrame, 
    period: str = "monthly",
    year: int = None,
    month: int = None
) -> float:
    """
    計算離職率
    
    Args:
        df: 員工資料 DataFrame
        period: 期間類型 - "weekly" 或 "monthly"
        year: 年份（僅 monthly 使用，預設為今年）
        month: 月份（僅 monthly 使用，預設為本月）
        
    Returns:
        離職率（百分比）
    """
    today = date.today()
    
    if period == "weekly":
        start_date = today - timedelta(days=today.weekday())
        end_date = today
        changes = get_period_changes(df, start_date, end_date)
    else:  # monthly（月基準為該月最後一日）
        if year is None:
            year = today.year
        if month is None:
            month = today.month
        
        changes = get_monthly_changes(df, year, month)
        
        start_date = date(year, month, 1)
        last_day = monthrange(year, month)[1]
        end_of_month = date(year, month, last_day)
        
        if year == today.year and month == today.month:
            end_date = min(today, end_of_month)
        else:
            end_date = end_of_month
    
    # 期初人數（使用月底人數 + 期間離職 - 期間入職 來反推期初）
    if period == "monthly":
        end_headcount = get_headcount_at_date(df, end_date)
        initial_headcount = end_headcount + changes['leaves'] - changes['hires']
    else:
        current_headcount = get_current_headcount(df)
        initial_headcount = current_headcount + changes['leaves'] - changes['hires']
    
    if initial_headcount <= 0:
        return 0.0
    
    turnover_rate = (changes['leaves'] / initial_headcount) * 100
    return round(turnover_rate, 2)


def get_department_stats(df: pd.DataFrame) -> pd.DataFrame:
    """
    取得各部門統計
    
    Args:
        df: 員工資料 DataFrame
        
    Returns:
        部門統計 DataFrame：department, count, direct_count, indirect_count
    """
    active = get_active_employees(df)
    
    # 基本人數統計
    dept_counts = active.groupby('department').size().reset_index(name='count')
    
    # 直接人員統計
    direct_counts = active[active['labor_type'] == '直接'].groupby('department').size().reset_index(name='direct_count')
    
    # 間接人員統計
    indirect_counts = active[active['labor_type'] == '間接'].groupby('department').size().reset_index(name='indirect_count')
    
    # 合併
    result = dept_counts.merge(direct_counts, on='department', how='left')
    result = result.merge(indirect_counts, on='department', how='left')
    
    # 填充空值
    result['direct_count'] = result['direct_count'].fillna(0).astype(int)
    result['indirect_count'] = result['indirect_count'].fillna(0).astype(int)
    
    return result.sort_values('count', ascending=False)


def get_business_unit_stats(df: pd.DataFrame) -> pd.DataFrame:
    """
    取得各事業部（部別）統計
    
    Args:
        df: 員工資料 DataFrame
        
    Returns:
        事業部統計 DataFrame：department, count, direct_count, indirect_count
    """
    active = get_active_employees(df)
    
    # 使用 business_unit 欄位，如沒有則用 department
    group_col = 'business_unit' if 'business_unit' in active.columns else 'department'
    
    # 基本人數統計
    bu_counts = active.groupby(group_col).size().reset_index(name='count')
    bu_counts = bu_counts.rename(columns={group_col: 'department'})
    
    # 直接人員統計
    direct_counts = active[active['labor_type'] == '直接'].groupby(group_col).size().reset_index(name='direct_count')
    direct_counts = direct_counts.rename(columns={group_col: 'department'})
    
    # 間接人員統計
    indirect_counts = active[active['labor_type'] == '間接'].groupby(group_col).size().reset_index(name='indirect_count')
    indirect_counts = indirect_counts.rename(columns={group_col: 'department'})
    
    # 合併
    result = bu_counts.merge(direct_counts, on='department', how='left')
    result = result.merge(indirect_counts, on='department', how='left')
    
    # 填充空值
    result['direct_count'] = result['direct_count'].fillna(0).astype(int)
    result['indirect_count'] = result['indirect_count'].fillna(0).astype(int)
    
    return result.sort_values('count', ascending=False)


def get_department_monthly_leaves(df: pd.DataFrame) -> pd.DataFrame:
    """
    取得各部門本月離職人數
    
    Args:
        df: 員工資料 DataFrame
        
    Returns:
        部門離職統計 DataFrame
    """
    monthly_changes = get_monthly_changes(df)
    leave_list = monthly_changes['leave_list']
    
    if len(leave_list) == 0:
        return pd.DataFrame(columns=['department', 'leave_count'])
    
    dept_leaves = leave_list.groupby('department').size().reset_index(name='leave_count')
    return dept_leaves.sort_values('leave_count', ascending=False)


def get_position_stats(df: pd.DataFrame) -> pd.DataFrame:
    """
    取得各職務統計
    
    Args:
        df: 員工資料 DataFrame
        
    Returns:
        職務統計 DataFrame：position, count
    """
    active = get_active_employees(df)
    position_counts = active.groupby('position').size().reset_index(name='count')
    return position_counts.sort_values('count', ascending=False)


def get_labor_type_stats(df: pd.DataFrame) -> pd.DataFrame:
    """
    取得直間接統計
    
    Args:
        df: 員工資料 DataFrame
        
    Returns:
        直間接統計 DataFrame：labor_type, count
    """
    active = get_active_employees(df)
    labor_counts = active.groupby('labor_type').size().reset_index(name='count')
    return labor_counts


def get_employment_type_stats(df: pd.DataFrame) -> pd.DataFrame:
    """
    取得員工狀態統計（原僱用類型）
    
    Args:
        df: 員工資料 DataFrame
        
    Returns:
        員工狀態統計 DataFrame：employee_status, count
    """
    active = get_active_employees(df)
    
    # 使用 employee_status_original 欄位（如果存在）
    status_col = 'employee_status_original' if 'employee_status_original' in active.columns else 'employment_type'
    
    if status_col not in active.columns:
        # 如果都不存在，返回空的統計
        return pd.DataFrame({'employee_status': [], 'count': []})
    
    emp_counts = active.groupby(status_col).size().reset_index(name='count')
    emp_counts.columns = ['employee_status', 'count']
    return emp_counts


def get_headcount_trend(
    df: pd.DataFrame, 
    weeks: int = 12
) -> pd.DataFrame:
    """
    計算人數趨勢（過去N週）
    
    Args:
        df: 員工資料 DataFrame
        weeks: 週數
        
    Returns:
        趨勢 DataFrame：date, headcount
    """
    df = df.copy()
    df['hire_date'] = pd.to_datetime(df['hire_date']).dt.date
    df['leave_date'] = pd.to_datetime(df['leave_date']).dt.date
    
    today = date.today()
    trend_data = []
    
    for i in range(weeks, -1, -1):
        # 計算該週末的日期
        target_date = today - timedelta(weeks=i)
        
        # 計算該日期的在職人數
        # 在職條件：hire_date <= target_date AND (leave_date is null OR leave_date > target_date)
        active_count = len(df[
            (df['hire_date'] <= target_date) & 
            ((df['leave_date'].isna()) | (df['leave_date'] > target_date))
        ])
        
        trend_data.append({
            'date': target_date,
            'headcount': active_count
        })
    
    return pd.DataFrame(trend_data)


def get_semiannual_trend(
    df: pd.DataFrame,
    start_year: int = 2024
) -> pd.DataFrame:
    """
    計算半年度人數趨勢（每年1月底和7月底）
    
    Args:
        df: 員工資料 DataFrame（應已過濾排除項目）
        start_year: 起始年份
        
    Returns:
        趨勢 DataFrame：date, headcount, label
    """
    df = df.copy()
    df['hire_date'] = pd.to_datetime(df['hire_date']).dt.date
    df['leave_date'] = pd.to_datetime(df['leave_date']).dt.date
    
    today = date.today()
    trend_data = []
    
    current_year = start_year
    while True:
        for month in [1, 7]:
            # 取得該月月底日期
            last_day = monthrange(current_year, month)[1]
            target_date = date(current_year, month, last_day)
            
            # 如果超過今天，就到此為止（但仍計算現在的數據）
            if target_date > today:
                # 加入今天的數據作為最後一筆
                active_count = len(df[
                    (df['hire_date'] <= today) & 
                    ((df['leave_date'].isna()) | (df['leave_date'] > today))
                ])
                trend_data.append({
                    'date': today,
                    'headcount': active_count,
                    'label': f"{today.year}年{today.month}月"
                })
                return pd.DataFrame(trend_data)
            
            # 計算該日期的在職人數
            active_count = len(df[
                (df['hire_date'] <= target_date) & 
                ((df['leave_date'].isna()) | (df['leave_date'] > target_date))
            ])
            
            trend_data.append({
                'date': target_date,
                'headcount': active_count,
                'label': f"{current_year}年{month}月"
            })
        
        current_year += 1
        
        # 防止無限迴圈
        if current_year > today.year + 1:
            break
    
    return pd.DataFrame(trend_data)


def generate_snapshot_data(df: pd.DataFrame) -> Dict[str, Any]:
    """
    產生快照資料
    
    Args:
        df: 員工資料 DataFrame
        
    Returns:
        快照資料字典
    """
    today = date.today()
    weekly_changes = get_weekly_changes(df)
    
    return {
        'snapshot_date': today,
        'total_headcount': get_current_headcount(df),
        'by_department': get_department_stats(df).to_dict('records'),
        'by_position': get_position_stats(df).to_dict('records'),
        'by_employment_type': get_employment_type_stats(df).to_dict('records'),
        'by_labor_type': get_labor_type_stats(df).to_dict('records'),
        'weekly_hires': weekly_changes['hires'],
        'weekly_leaves': weekly_changes['leaves'],
        'monthly_turnover_rate': get_turnover_rate(df, 'monthly')
    }
