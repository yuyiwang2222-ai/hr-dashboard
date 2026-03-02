# 大豐環保人力變化視覺化系統 - 模組初始化
# ============================================

from .data_loader import load_employees, load_departments, load_config
from .metrics import (
    get_current_headcount,
    get_period_changes,
    get_turnover_rate,
    get_department_stats,
    get_position_stats,
    get_labor_type_stats,
    get_employment_type_stats
)
from .charts import (
    create_trend_chart,
    create_department_chart,
    create_position_pie,
    create_employment_type_pie,
    create_labor_type_pie
)
from .alerts import check_all_alerts, Alert

__all__ = [
    # data_loader
    'load_employees',
    'load_departments', 
    'load_config',
    # metrics
    'get_current_headcount',
    'get_period_changes',
    'get_turnover_rate',
    'get_department_stats',
    'get_position_stats',
    'get_labor_type_stats',
    'get_employment_type_stats',
    # charts
    'create_trend_chart',
    'create_department_chart',
    'create_position_pie',
    'create_employment_type_pie',
    'create_labor_type_pie',
    # alerts
    'check_all_alerts',
    'Alert'
]
