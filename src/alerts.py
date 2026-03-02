"""
大豐環保人力儀表板 - 異常警報模組
============================================
負責偵測異常狀況並產生警報
"""

import pandas as pd
from datetime import date, datetime
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
from enum import Enum

from .metrics import (
    get_turnover_rate,
    get_department_monthly_leaves,
    get_current_headcount,
    get_weekly_changes
)


class AlertLevel(Enum):
    """警報等級"""
    INFO = "info"       # 資訊
    WARNING = "warning"  # 警告（黃燈）
    CRITICAL = "critical"  # 危急（紅燈）


@dataclass
class Alert:
    """警報資料類別"""
    level: AlertLevel
    title: str
    message: str
    metric_name: str
    current_value: float
    threshold: float
    timestamp: datetime = None
    
    def __post_init__(self):
        if self.timestamp is None:
            self.timestamp = datetime.now()
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'level': self.level.value,
            'title': self.title,
            'message': self.message,
            'metric_name': self.metric_name,
            'current_value': self.current_value,
            'threshold': self.threshold,
            'timestamp': self.timestamp.isoformat()
        }


def check_turnover_rate_alert(
    df: pd.DataFrame,
    config: Dict[str, Any]
) -> Optional[Alert]:
    """
    檢查離職率警報
    
    Args:
        df: 員工資料 DataFrame
        config: 設定字典
        
    Returns:
        Alert 物件或 None
    """
    turnover_rate = get_turnover_rate(df, 'monthly')
    thresholds = config.get('alerts', {}).get('monthly_turnover_rate', {})
    
    critical_threshold = thresholds.get('critical', 10)
    warning_threshold = thresholds.get('warning', 5)
    
    if turnover_rate >= critical_threshold:
        return Alert(
            level=AlertLevel.CRITICAL,
            title='🔴 離職率過高',
            message=f'本月離職率 {turnover_rate:.1f}% 超過危急閾值 {critical_threshold}%',
            metric_name='monthly_turnover_rate',
            current_value=turnover_rate,
            threshold=critical_threshold
        )
    elif turnover_rate >= warning_threshold:
        return Alert(
            level=AlertLevel.WARNING,
            title='🟡 離職率偏高',
            message=f'本月離職率 {turnover_rate:.1f}% 超過警告閾值 {warning_threshold}%',
            metric_name='monthly_turnover_rate',
            current_value=turnover_rate,
            threshold=warning_threshold
        )
    
    return None


def check_department_leaves_alert(
    df: pd.DataFrame,
    config: Dict[str, Any],
    departments_df: pd.DataFrame = None
) -> List[Alert]:
    """
    檢查部門離職人數警報
    
    Args:
        df: 員工資料 DataFrame
        config: 設定字典
        departments_df: 部門對照表
        
    Returns:
        Alert 物件列表
    """
    alerts = []
    dept_leaves = get_department_monthly_leaves(df)
    threshold = config.get('alerts', {}).get('department_monthly_leaves', {}).get('warning', 3)
    
    for _, row in dept_leaves.iterrows():
        if row['leave_count'] >= threshold:
            dept_code = row['department']
            dept_name = dept_code
            
            # 取得部門名稱
            if departments_df is not None:
                match = departments_df[departments_df['code'] == dept_code]
                if len(match) > 0:
                    dept_name = match.iloc[0]['name']
            
            alerts.append(Alert(
                level=AlertLevel.WARNING,
                title=f'🟡 {dept_name}離職人數偏高',
                message=f'{dept_name}本月離職 {row["leave_count"]} 人，超過閾值 {threshold} 人',
                metric_name='department_monthly_leaves',
                current_value=row['leave_count'],
                threshold=threshold
            ))
    
    return alerts


def check_headcount_change_alert(
    df: pd.DataFrame,
    config: Dict[str, Any]
) -> Optional[Alert]:
    """
    檢查人數變化警報
    
    Args:
        df: 員工資料 DataFrame
        config: 設定字典
        
    Returns:
        Alert 物件或 None
    """
    weekly_changes = get_weekly_changes(df)
    current_headcount = get_current_headcount(df)
    threshold = config.get('alerts', {}).get('weekly_headcount_change', {}).get('warning', 5)
    
    # 計算週變化率
    if current_headcount > 0:
        net_change = weekly_changes['net_change']
        # 期初人數 = 目前人數 - 淨變化
        initial_headcount = current_headcount - net_change
        if initial_headcount > 0:
            change_rate = abs(net_change / initial_headcount) * 100
            
            if change_rate >= threshold:
                direction = '增加' if net_change > 0 else '減少'
                return Alert(
                    level=AlertLevel.WARNING,
                    title='🟡 人力大幅變動',
                    message=f'本週人數{direction} {abs(net_change)} 人（變化率 {change_rate:.1f}%），超過閾值 {threshold}%',
                    metric_name='weekly_headcount_change',
                    current_value=change_rate,
                    threshold=threshold
                )
    
    return None


def check_all_alerts(
    df: pd.DataFrame,
    config: Dict[str, Any],
    departments_df: pd.DataFrame = None
) -> List[Alert]:
    """
    檢查所有警報
    
    Args:
        df: 員工資料 DataFrame
        config: 設定字典
        departments_df: 部門對照表
        
    Returns:
        所有警報列表
    """
    alerts = []
    
    # 檢查離職率
    turnover_alert = check_turnover_rate_alert(df, config)
    if turnover_alert:
        alerts.append(turnover_alert)
    
    # 檢查部門離職
    dept_alerts = check_department_leaves_alert(df, config, departments_df)
    alerts.extend(dept_alerts)
    
    # 檢查人數變化
    headcount_alert = check_headcount_change_alert(df, config)
    if headcount_alert:
        alerts.append(headcount_alert)
    
    # 依嚴重程度排序（危急優先）
    level_order = {AlertLevel.CRITICAL: 0, AlertLevel.WARNING: 1, AlertLevel.INFO: 2}
    alerts.sort(key=lambda a: level_order.get(a.level, 99))
    
    return alerts


def format_alert_message(alert: Alert) -> str:
    """
    格式化警報訊息（用於 Email 或通知）
    
    Args:
        alert: Alert 物件
        
    Returns:
        格式化後的訊息字串
    """
    level_emoji = {
        AlertLevel.CRITICAL: '🔴',
        AlertLevel.WARNING: '🟡',
        AlertLevel.INFO: 'ℹ️'
    }
    
    emoji = level_emoji.get(alert.level, '')
    
    return f"""
{emoji} {alert.title}
━━━━━━━━━━━━━━━━━━━━━━
{alert.message}

📊 當前數值：{alert.current_value}
⚠️ 警報閾值：{alert.threshold}
🕐 時間：{alert.timestamp.strftime('%Y-%m-%d %H:%M')}
"""


def get_alerts_summary(alerts: List[Alert]) -> str:
    """
    取得警報摘要
    
    Args:
        alerts: 警報列表
        
    Returns:
        摘要字串
    """
    if not alerts:
        return "✅ 目前無異常警報"
    
    critical_count = sum(1 for a in alerts if a.level == AlertLevel.CRITICAL)
    warning_count = sum(1 for a in alerts if a.level == AlertLevel.WARNING)
    
    summary = f"共 {len(alerts)} 項警報"
    if critical_count > 0:
        summary += f"（🔴 危急 {critical_count}）"
    if warning_count > 0:
        summary += f"（🟡 警告 {warning_count}）"
    
    return summary
