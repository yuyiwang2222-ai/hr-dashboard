"""
大豐環保人力儀表板 - 圖表產生模組
============================================
負責產生各種 Plotly 圖表
"""

import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from typing import Dict, Optional, List

# 預設色彩配置
DEFAULT_COLORS = {
    'primary': '#1E88E5',
    'success': '#4CAF50',
    'warning': '#FFC107',
    'danger': '#F44336',
    'neutral': '#9E9E9E',
    'background': '#FAFAFA'
}

# 部門色彩對照（固定配色以便識別）
DEPT_COLORS = {
    'GM': '#5C6BC0',    # 總經理室 - 靛藍
    'BD': '#42A5F5',    # 業務開發部 - 藍
    'OPS': '#26A69A',   # 營運管理部 - 青
    'REC': '#66BB6A',   # 再生處理部 - 綠
    'INT': '#FFA726',   # 國際事業部 - 橙
    'EC': '#EC407A',    # 電子商務部 - 粉紅
    'MKT': '#AB47BC',   # 行銷部 - 紫
    'RD': '#7E57C2',    # 研發部 - 深紫
    'FIN': '#78909C',   # 財務部 - 藍灰
    'HR': '#8D6E63',    # 人資部 - 棕
    'IT': '#26C6DA'     # 資訊部 - 青藍
}


def create_trend_chart(
    trend_df: pd.DataFrame,
    colors: Optional[Dict[str, str]] = None,
    title: str = '人數趨勢（過去12週）'
) -> go.Figure:
    """
    建立人數趨勢折線圖
    
    Args:
        trend_df: 趨勢資料 DataFrame（包含 date, headcount 欄位，可選 label 欄位）
        colors: 色彩配置字典
        title: 圖表標題
        
    Returns:
        Plotly Figure 物件
    """
    colors = colors or DEFAULT_COLORS
    
    fig = go.Figure()
    
    # 使用 label 欄位作為顯示文字（如果有的話）
    x_values = trend_df['label'] if 'label' in trend_df.columns else trend_df['date']
    hover_template = '%{x}<br>人數: %{y}<extra></extra>'
    
    fig.add_trace(go.Scatter(
        x=x_values,
        y=trend_df['headcount'],
        mode='lines+markers',
        name='人數',
        line=dict(color=colors['primary'], width=3),
        marker=dict(size=8),
        hovertemplate=hover_template
    ))
    
    fig.update_layout(
        title=title,
        xaxis_title='日期',
        yaxis_title='人數',
        hovermode='x unified',
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(family='Microsoft JhengHei, Arial', size=12),
        margin=dict(l=40, r=40, t=60, b=40)
    )
    
    # 設定 Y 軸從 0 開始，並留一些空間
    if len(trend_df) > 0:
        max_val = trend_df['headcount'].max()
        fig.update_yaxes(range=[0, max_val * 1.1])
    
    fig.update_xaxes(showgrid=True, gridcolor='#E0E0E0')
    fig.update_yaxes(showgrid=True, gridcolor='#E0E0E0')
    
    return fig


def create_department_chart(
    dept_df: pd.DataFrame,
    departments_df: pd.DataFrame,
    show_labor_type: str = 'all',
    colors: Optional[Dict[str, str]] = None
) -> go.Figure:
    """
    建立部門人數長條圖
    
    Args:
        dept_df: 部門統計 DataFrame
        departments_df: 部門對照表
        show_labor_type: 顯示類型 - 'all', 'direct', 'indirect'
        colors: 色彩配置字典
        
    Returns:
        Plotly Figure 物件
    """
    colors = colors or DEFAULT_COLORS
    
    # 合併部門名稱
    merged = dept_df.merge(
        departments_df[['code', 'name']],
        left_on='department',
        right_on='code',
        how='left'
    )
    merged['dept_name'] = merged['name'].fillna(merged['department'])
    
    # 根據顯示類型選擇欄位
    if show_labor_type == 'direct':
        y_col = 'direct_count'
        title = '各部門直接人員數'
    elif show_labor_type == 'indirect':
        y_col = 'indirect_count'
        title = '各部門間接人員數'
    else:
        y_col = 'count'
        title = '各部門人數'
    
    # 排序
    merged = merged.sort_values(y_col, ascending=True)
    
    # 建立顏色列表
    bar_colors = [DEPT_COLORS.get(code, colors['primary']) for code in merged['department']]
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=merged[y_col],
        y=merged['dept_name'],
        orientation='h',
        marker_color=bar_colors,
        text=merged[y_col],
        textposition='outside',
        hovertemplate='%{y}<br>人數: %{x}<extra></extra>'
    ))
    
    fig.update_layout(
        title=title,
        xaxis_title='人數',
        yaxis_title='',
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(family='Microsoft JhengHei, Arial', size=12),
        margin=dict(l=120, r=40, t=60, b=40),
        showlegend=False
    )
    
    fig.update_xaxes(showgrid=True, gridcolor='#E0E0E0')
    
    return fig


def create_department_stacked_chart(
    dept_df: pd.DataFrame,
    departments_df: pd.DataFrame,
    colors: Optional[Dict[str, str]] = None
) -> go.Figure:
    """
    建立部門直間接堆疊長條圖
    
    Args:
        dept_df: 部門統計 DataFrame
        departments_df: 部門對照表
        colors: 色彩配置字典
        
    Returns:
        Plotly Figure 物件
    """
    colors = colors or DEFAULT_COLORS
    
    # 合併部門名稱
    merged = dept_df.merge(
        departments_df[['code', 'name']],
        left_on='department',
        right_on='code',
        how='left'
    )
    merged['dept_name'] = merged['name'].fillna(merged['department'])
    merged = merged.sort_values('count', ascending=True)
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        name='直接',
        x=merged['direct_count'],
        y=merged['dept_name'],
        orientation='h',
        marker_color=colors['primary'],
        hovertemplate='%{y}<br>直接人員: %{x}<extra></extra>'
    ))
    
    fig.add_trace(go.Bar(
        name='間接',
        x=merged['indirect_count'],
        y=merged['dept_name'],
        orientation='h',
        marker_color=colors['success'],
        hovertemplate='%{y}<br>間接人員: %{x}<extra></extra>'
    ))
    
    fig.update_layout(
        title='各部門直接/間接人員分布',
        xaxis_title='人數',
        yaxis_title='',
        barmode='stack',
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(family='Microsoft JhengHei, Arial', size=12),
        margin=dict(l=120, r=40, t=60, b=40),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
    )
    
    fig.update_xaxes(showgrid=True, gridcolor='#E0E0E0')
    
    return fig


def create_position_pie(
    position_df: pd.DataFrame,
    colors: Optional[Dict[str, str]] = None
) -> go.Figure:
    """
    建立職務分布圓餅圖
    
    Args:
        position_df: 職務統計 DataFrame
        colors: 色彩配置字典
        
    Returns:
        Plotly Figure 物件
    """
    # 計算百分比並建立標籤
    total = position_df['count'].sum()
    position_df = position_df.copy()
    position_df['percentage'] = (position_df['count'] / total * 100).round(1)
    position_df['label'] = position_df.apply(
        lambda row: f"{row['position']}<br>{row['count']}人 ({row['percentage']}%)", axis=1
    )
    
    fig = px.pie(
        position_df,
        values='count',
        names='position',
        title='職務分布',
        hole=0.4  # 甜甜圈圖
    )
    
    fig.update_traces(
        textposition='inside',
        textinfo='percent+label',
        texttemplate='%{label}<br>%{value}人',
        hovertemplate='%{label}<br>人數: %{value}<br>佔比: %{percent}<extra></extra>'
    )
    
    fig.update_layout(
        font=dict(family='Microsoft JhengHei, Arial', size=12),
        margin=dict(l=20, r=20, t=60, b=20),
        showlegend=True,
        legend=dict(orientation='v', yanchor='middle', y=0.5, xanchor='left', x=1.02)
    )
    
    return fig


def create_employment_type_pie(
    emp_type_df: pd.DataFrame,
    colors: Optional[Dict[str, str]] = None
) -> go.Figure:
    """
    建立員工狀態圓餅圖
    
    Args:
        emp_type_df: 員工狀態統計 DataFrame（欄位：employee_status, count）
        colors: 色彩配置字典
        
    Returns:
        Plotly Figure 物件
    """
    colors = colors or DEFAULT_COLORS
    
    if len(emp_type_df) == 0:
        # 無資料時返回空圖
        fig = go.Figure()
        fig.update_layout(title='員工狀態分布（無資料）')
        return fig
    
    # 判斷使用哪個欄位
    label_col = 'employee_status' if 'employee_status' in emp_type_df.columns else 'employment_type'
    
    # 自訂顏色
    color_map = {
        '正式員工': colors['primary'],
        '試用員工': colors['success'],
        '工讀生': colors['warning'],
        '約聘': '#FF9800',
        '實習生': '#9C27B0',
        '正職': colors['primary'],
        '約聘': colors['warning'],
        '實習': colors['success']
    }
    pie_colors = [color_map.get(str(t), colors['neutral']) for t in emp_type_df[label_col]]
    
    fig = go.Figure(data=[go.Pie(
        labels=emp_type_df[label_col],
        values=emp_type_df['count'],
        hole=0.4,
        marker_colors=pie_colors
    )])
    
    fig.update_traces(
        textposition='inside',
        textinfo='percent+label',
        hovertemplate='%{label}<br>人數: %{value}<br>佔比: %{percent}<extra></extra>'
    )
    
    fig.update_layout(
        title='員工狀態分布',
        font=dict(family='Microsoft JhengHei, Arial', size=12),
        margin=dict(l=20, r=20, t=60, b=20),
        showlegend=True
    )
    
    return fig


def create_labor_type_pie(
    labor_df: pd.DataFrame,
    colors: Optional[Dict[str, str]] = None
) -> go.Figure:
    """
    建立直間接人員圓餅圖
    
    Args:
        labor_df: 直間接統計 DataFrame
        colors: 色彩配置字典
        
    Returns:
        Plotly Figure 物件
    """
    colors = colors or DEFAULT_COLORS
    
    # 自訂顏色
    color_map = {
        '直接': colors['primary'],
        '間接': colors['success']
    }
    pie_colors = [color_map.get(t, colors['neutral']) for t in labor_df['labor_type']]
    
    fig = go.Figure(data=[go.Pie(
        labels=labor_df['labor_type'],
        values=labor_df['count'],
        hole=0.4,
        marker_colors=pie_colors
    )])
    
    fig.update_traces(
        textposition='inside',
        textinfo='percent+label',
        hovertemplate='%{label}<br>人數: %{value}<br>佔比: %{percent}<extra></extra>'
    )
    
    fig.update_layout(
        title='直接/間接人員分布',
        font=dict(family='Microsoft JhengHei, Arial', size=12),
        margin=dict(l=20, r=20, t=60, b=20),
        showlegend=True
    )
    
    return fig


def create_kpi_indicator(
    value: float,
    title: str,
    delta: Optional[float] = None,
    is_percentage: bool = False,
    colors: Optional[Dict[str, str]] = None
) -> go.Figure:
    """
    建立 KPI 指標卡片
    
    Args:
        value: 數值
        title: 標題
        delta: 變化量（正為增加，負為減少）
        is_percentage: 是否為百分比
        colors: 色彩配置
        
    Returns:
        Plotly Figure 物件
    """
    colors = colors or DEFAULT_COLORS
    
    # 格式化數值
    if is_percentage:
        number_format = '.1%' if value < 1 else '.2f'
        suffix = '%' if value < 1 else '%'
    else:
        number_format = None
        suffix = ''
    
    fig = go.Figure()
    
    fig.add_trace(go.Indicator(
        mode='number+delta' if delta is not None else 'number',
        value=value,
        title={'text': title, 'font': {'size': 16}},
        number={'font': {'size': 36}, 'suffix': suffix if not is_percentage else '%'},
        delta={'reference': value - delta, 'relative': False} if delta is not None else None,
        domain={'x': [0, 1], 'y': [0, 1]}
    ))
    
    fig.update_layout(
        height=150,
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor='white'
    )
    
    return fig
