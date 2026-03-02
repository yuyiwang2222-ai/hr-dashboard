"""
大豐環保人力儀表板 - 資料載入模組
============================================
負責讀取 Excel 資料檔案與設定檔
支援從「數據資料夾」讀取外部資料並做欄位對應
支援事業部對應表、工讀生彈性篩選
"""

import pandas as pd
import yaml
import os
import glob
from datetime import datetime, date
from typing import Optional, Dict, Any, List


# 全域變數：事業部對應表快取
_business_unit_mapping_cache = None
# 全域變數：職務類別對應表快取
_job_category_mapping_cache = None


def clear_cache():
    """
    清除所有快取
    用於重新載入資料時確保取得最新資料
    """
    global _business_unit_mapping_cache, _job_category_mapping_cache
    _business_unit_mapping_cache = None
    _job_category_mapping_cache = None
    print("✅ 已清除 data_loader 快取")


def load_config(config_path: str = "config.yaml") -> Dict[str, Any]:
    """
    載入 YAML 設定檔
    
    Args:
        config_path: 設定檔路徑
        
    Returns:
        設定字典
    """
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        return config
    except FileNotFoundError:
        print(f"⚠️ 找不到設定檔：{config_path}，使用預設設定")
        return get_default_config()
    except Exception as e:
        print(f"⚠️ 讀取設定檔錯誤：{e}，使用預設設定")
        return get_default_config()


def get_default_config() -> Dict[str, Any]:
    """取得預設設定"""
    return {
        'system': {
            'app_title': '大豐環保人力儀表板',
            'data_path': './data',
            'snapshot_path': './data/snapshots'
        },
        'alerts': {
            'monthly_turnover_rate': {'warning': 5, 'critical': 10},
            'department_monthly_leaves': {'warning': 3},
            'weekly_headcount_change': {'warning': 5}
        },
        'display': {
            'trend_weeks': 12,
            'colors': {
                'primary': '#1E88E5',
                'success': '#4CAF50',
                'warning': '#FFC107',
                'danger': '#F44336',
                'neutral': '#9E9E9E',
                'background': '#FAFAFA'
            }
        },
        'employee_status': {
            'active_statuses': ['正式員工', '試用員工', '工讀生', '約聘', '實習生'],
            'part_time_statuses': ['工讀生', '實習生'],
            'inactive_statuses': ['退休員工', '離職員工', '留職停薪']
        }
    }


def load_business_unit_mapping(external_path: str, config: Dict[str, Any]) -> Dict[str, str]:
    """
    載入事業部對應表
    
    Args:
        external_path: 外部資料夾路徑
        config: 設定字典
        
    Returns:
        部門名稱 → 事業部 對應字典
    """
    global _business_unit_mapping_cache
    
    # 檢查是否有快取
    if _business_unit_mapping_cache is not None:
        return _business_unit_mapping_cache
    
    bu_config = config.get('business_unit_mapping', {})
    if not bu_config.get('enabled', False):
        return {}
    
    file_name = bu_config.get('file_name', '事業部對應表.xlsx')
    file_path = os.path.join(external_path, file_name)
    
    if not os.path.exists(file_path):
        print(f"⚠️ 找不到事業部對應表：{file_path}")
        return {}
    
    try:
        # 讀取對應表（指定工作表名稱）
        sheet_name = bu_config.get('department_sheet', '處別')
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        dept_col = bu_config.get('department_column', '部門名稱')
        bu_col = bu_config.get('business_unit_column', '部別')
        
        if dept_col not in df.columns or bu_col not in df.columns:
            print(f"⚠️ 事業部對應表缺少必要欄位：{dept_col} 或 {bu_col}")
            print(f"   可用欄位：{list(df.columns)}")
            return {}
        
        # 建立對應字典
        mapping = dict(zip(df[dept_col].astype(str).str.strip(), df[bu_col].astype(str).str.strip()))
        
        # 快取
        _business_unit_mapping_cache = mapping
        print(f"✅ 載入事業部對應表：{len(mapping)} 個部門對應")
        
        return mapping
        
    except Exception as e:
        print(f"⚠️ 讀取事業部對應表錯誤：{e}")
        return {}


def load_job_category_mapping(external_path: str, config: Dict[str, Any]) -> Dict[str, str]:
    """
    載入職務類別對應表
    
    Args:
        external_path: 外部資料夾路徑
        config: 設定字典
        
    Returns:
        所屬職務 → 職務類別 對應字典
    """
    global _job_category_mapping_cache
    
    # 檢查是否有快取
    if _job_category_mapping_cache is not None:
        return _job_category_mapping_cache
    
    jc_config = config.get('job_category_mapping', {})
    if not jc_config.get('enabled', False):
        return {}
    
    file_name = jc_config.get('file_name', '事業部對應表.xlsx')
    file_path = os.path.join(external_path, file_name)
    
    if not os.path.exists(file_path):
        print(f"⚠️ 找不到職務對應表：{file_path}")
        return {}
    
    try:
        # 讀取對應表（指定工作表名稱）
        sheet_name = jc_config.get('job_sheet', '職務')
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        job_title_col = jc_config.get('job_title_column', '所屬職務')
        job_category_col = jc_config.get('job_category_column', '職務類別')
        
        if job_title_col not in df.columns or job_category_col not in df.columns:
            print(f"⚠️ 職務對應表缺少必要欄位：{job_title_col} 或 {job_category_col}")
            print(f"   可用欄位：{list(df.columns)}")
            return {}
        
        # 建立對應字典
        mapping = dict(zip(
            df[job_title_col].dropna().astype(str).str.strip(), 
            df[job_category_col].dropna().astype(str).str.strip()
        ))
        
        # 快取
        _job_category_mapping_cache = mapping
        print(f"✅ 載入職務類別對應表：{len(mapping)} 個職務對應")
        
        return mapping
        
    except Exception as e:
        print(f"⚠️ 讀取職務對應表錯誤：{e}")
        return {}


def load_employees(file_path: str = "data/employees.xlsx", config: Dict[str, Any] = None) -> pd.DataFrame:
    """
    載入員工主檔
    優先從外部資料夾（數據資料夾）讀取，若無則從預設路徑讀取
    
    Args:
        file_path: 預設 Excel 檔案路徑
        config: 設定字典（用於取得外部資料路徑和欄位對應）
        
    Returns:
        員工資料 DataFrame
    """
    # 載入設定
    if config is None:
        config = load_config()
    
    system_config = config.get('system', {})
    
    # 檢查是否使用外部資料
    if system_config.get('use_external_data', False):
        external_path = system_config.get('external_data_path', './數據資料夾')
        # 只有當外部資料夾存在時才嘗試讀取
        if os.path.exists(external_path):
            external_df = load_external_employees(external_path, config)
            if len(external_df) > 0:
                print(f"✅ 從外部資料夾載入 {len(external_df)} 筆員工資料")
                return external_df
            else:
                print("⚠️ 外部資料夾無有效資料，改用預設資料")
        else:
            print(f"📂 外部資料夾不存在，使用 data/ 資料夾")
    
    # 從預設路徑載入（data/employees.xlsx）
    try:
        df = pd.read_excel(file_path, sheet_name=0)
        
        # 檢查是否需要做欄位對應（如果有中文欄位名稱）
        if '員工自然編號' in df.columns or '中文名' in df.columns or '到職日期' in df.columns:
            # 原始資料格式，需要做欄位對應
            column_mapping = config.get('column_mapping', {})
            status_config = config.get('employee_status', {})
            jc_config = config.get('job_category_mapping', {})
            
            # 嘗試從 data/ 資料夾讀取對應表
            data_path = os.path.dirname(file_path) if file_path else './data'
            bu_mapping = {}
            jc_mapping = {}
            
            # 檢查是否有對應表檔案在 data/ 資料夾
            dept_file = os.path.join(data_path, 'departments.xlsx')
            if os.path.exists(dept_file):
                try:
                    bu_config = config.get('business_unit_mapping', {})
                    sheet_name = bu_config.get('department_sheet', '處別')
                    dept_col = bu_config.get('department_column', '部門名稱')
                    bu_col = bu_config.get('business_unit_column', '部別')
                    
                    dept_df = pd.read_excel(dept_file, sheet_name=sheet_name)
                    if dept_col in dept_df.columns and bu_col in dept_df.columns:
                        bu_mapping = dict(zip(dept_df[dept_col].astype(str).str.strip(), 
                                            dept_df[bu_col].astype(str).str.strip()))
                        print(f"✅ 載入事業部對應表：{len(bu_mapping)} 個部門對應")
                    
                    # 職務類別對應
                    jc_sheet = jc_config.get('job_sheet', '職務')
                    job_title_col = jc_config.get('job_title_column', '所屬職務')
                    job_category_col = jc_config.get('job_category_column', '職務類別')
                    
                    jc_df = pd.read_excel(dept_file, sheet_name=jc_sheet)
                    if job_title_col in jc_df.columns and job_category_col in jc_df.columns:
                        jc_mapping = dict(zip(
                            jc_df[job_title_col].dropna().astype(str).str.strip(),
                            jc_df[job_category_col].dropna().astype(str).str.strip()
                        ))
                        print(f"✅ 載入職務類別對應表：{len(jc_mapping)} 個職務對應")
                except Exception as e:
                    print(f"⚠️ 讀取對應表錯誤：{e}")
            
            df = apply_column_mapping(df, column_mapping, bu_mapping, jc_mapping, status_config, jc_config)
            
            # 排除特定公司的員工
            excluded_companies = config.get('system', {}).get('excluded_companies', [])
            if excluded_companies and 'company' in df.columns:
                original_count = len(df)
                df = df[~df['company'].isin(excluded_companies)]
                excluded_count = original_count - len(df)
                if excluded_count > 0:
                    print(f"  ✅ 已排除 {excluded_count} 位員工（所屬公司：{', '.join(excluded_companies)}）")
            
            # 排除特定員工
            excluded_employees = config.get('system', {}).get('excluded_employees', [])
            if excluded_employees and 'name' in df.columns:
                original_count = len(df)
                df = df[~df['name'].isin(excluded_employees)]
                excluded_count = original_count - len(df)
                if excluded_count > 0:
                    print(f"  ✅ 已排除 {excluded_count} 位特定員工（{', '.join(excluded_employees)}）")
            
            print(f"✅ 從 data/ 資料夾載入 {len(df)} 筆員工資料")
            return df
        
        # 已經是標準格式
        # 確保日期欄位正確轉換
        if 'hire_date' in df.columns:
            df['hire_date'] = pd.to_datetime(df['hire_date']).dt.date
        if 'leave_date' in df.columns:
            df['leave_date'] = pd.to_datetime(df['leave_date']).dt.date
            
        # 填充空值
        df['status'] = df['status'].fillna('在職')
        df['employment_type'] = df['employment_type'].fillna('正職')
        df['labor_type'] = df['labor_type'].fillna('直接')
        
        return df
        
    except FileNotFoundError:
        print(f"⚠️ 找不到員工檔案：{file_path}")
        return pd.DataFrame()
    except Exception as e:
        print(f"⚠️ 讀取員工檔案錯誤：{e}")
        return pd.DataFrame()


def load_external_employees(external_path: str, config: Dict[str, Any]) -> pd.DataFrame:
    """
    從外部資料夾載入員工資料並做欄位對應
    
    Args:
        external_path: 外部資料夾路徑
        config: 設定字典
        
    Returns:
        處理後的員工 DataFrame
    """
    if not os.path.exists(external_path):
        print(f"⚠️ 外部資料夾不存在：{external_path}")
        return pd.DataFrame()
    
    # 尋找員工資料檔案（支援 xlsx 和 xls，排除暫存檔）
    xlsx_files = glob.glob(os.path.join(external_path, "員工*.xlsx"))
    xls_files = glob.glob(os.path.join(external_path, "員工*.xls"))
    all_files = xlsx_files + [f for f in xls_files if not f.endswith('.xlsx')]
    # 排除暫存檔（以 ~$ 開頭）
    all_files = [f for f in all_files if not os.path.basename(f).startswith('~$')]
    
    if not all_files:
        print(f"⚠️ 找不到員工資料檔案（員工*.xlsx）")
        return pd.DataFrame()
    
    # 找最新的檔案（依修改時間）
    latest_file = max(all_files, key=os.path.getmtime)
    print(f"📂 讀取員工資料：{os.path.basename(latest_file)}")
    
    try:
        # 讀取 Excel（支援 xls 和 xlsx）
        if latest_file.endswith('.xls') and not latest_file.endswith('.xlsx'):
            df = pd.read_excel(latest_file, sheet_name=0, engine='xlrd')
        else:
            df = pd.read_excel(latest_file, sheet_name=0)
        
        # 取得設定
        column_mapping = config.get('column_mapping', {})
        bu_mapping = load_business_unit_mapping(external_path, config)
        jc_mapping = load_job_category_mapping(external_path, config)
        status_config = config.get('employee_status', {})
        jc_config = config.get('job_category_mapping', {})
        
        # 做欄位對應
        df = apply_column_mapping(df, column_mapping, bu_mapping, jc_mapping, status_config, jc_config)
        
        # 排除特定公司的員工
        excluded_companies = config.get('system', {}).get('excluded_companies', [])
        if excluded_companies and 'company' in df.columns:
            before_count = len(df)
            df = df[~df['company'].isin(excluded_companies)]
            excluded_count = before_count - len(df)
            if excluded_count > 0:
                print(f"  ✅ 已排除 {excluded_count} 位員工（所屬公司：{', '.join(excluded_companies)}）")
        
        # 排除特定員工（例如身障人員）
        excluded_employees = config.get('system', {}).get('excluded_employees', [])
        if excluded_employees and 'name' in df.columns:
            before_count = len(df)
            df = df[~df['name'].isin(excluded_employees)]
            excluded_count = before_count - len(df)
            if excluded_count > 0:
                print(f"  ✅ 已排除 {excluded_count} 位特定員工（{', '.join(excluded_employees)}）")
        
        return df
        
    except Exception as e:
        print(f"⚠️ 讀取外部資料錯誤：{e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()


def apply_column_mapping(
    df: pd.DataFrame, 
    column_mapping: Dict[str, str],
    bu_mapping: Dict[str, str],
    jc_mapping: Dict[str, str],
    status_config: Dict[str, List[str]],
    jc_config: Dict[str, Any] = None
) -> pd.DataFrame:
    """
    套用欄位對應，將外部資料欄位轉換為系統欄位
    
    Args:
        df: 原始 DataFrame
        column_mapping: 欄位對應字典 {系統欄位: 外部欄位}
        bu_mapping: 部門名稱 → 事業部 對應字典
        jc_mapping: 所屬職務 → 職務類別 對應字典
        status_config: 員工狀態設定
        jc_config: 職務類別設定
        
    Returns:
        對應後的 DataFrame
    """
    if jc_config is None:
        jc_config = {}
    
    result = pd.DataFrame()
    
    # 對應每個欄位
    for system_col, external_col in column_mapping.items():
        if external_col in df.columns:
            result[system_col] = df[external_col]
        else:
            # 嘗試模糊匹配（欄位名稱包含關鍵字）
            matched = False
            for col in df.columns:
                if external_col in str(col) or str(col) in external_col:
                    result[system_col] = df[col]
                    print(f"  ⚡ 模糊匹配：{col} → {system_col}")
                    matched = True
                    break
            
            if not matched:
                print(f"  ⚠️ 找不到欄位：{external_col}（對應 {system_col}）")
    
    # 保留原始部門名稱
    if 'department' in result.columns:
        result['department_original'] = result['department'].copy()
    
    # 套用事業部對應（部門名稱 → 事業部）
    if 'department' in result.columns and bu_mapping:
        result['business_unit'] = result['department'].apply(
            lambda x: bu_mapping.get(str(x).strip(), '未分類') if pd.notna(x) else '未分類'
        )
        # 用事業部取代原本的 department（儀表板以事業部為主）
        result['department'] = result['business_unit']
        print(f"  ✅ 已套用事業部對應")
    
    # 套用職務類別對應（所屬職務 → 職務類別）
    if 'job_title' in result.columns and jc_mapping:
        result['job_title_original'] = result['job_title'].copy()
        result['job_category'] = result['job_title'].apply(
            lambda x: jc_mapping.get(str(x).strip(), '其他') if pd.notna(x) else '其他'
        )
        # 用職務類別取代原本的 position（儀表板顯示職務類別）
        result['position'] = result['job_category']
        print(f"  ✅ 已套用職務類別對應")
    
    # 套用職務類別特例覆蓋（根據員工編號）
    jc_overrides = jc_config.get('overrides', {})
    if jc_overrides and 'employee_id' in result.columns:
        override_count = 0
        for emp_id, new_category in jc_overrides.items():
            emp_id_str = str(emp_id).strip()
            mask = result['employee_id'].astype(str).str.strip() == emp_id_str
            if mask.any():
                result.loc[mask, 'job_category'] = new_category
                result.loc[mask, 'position'] = new_category
                override_count += mask.sum()
        if override_count > 0:
            print(f"  ✅ 已套用 {override_count} 筆職務類別特例覆蓋")
    
    # 特殊處理：人資部、資訊部、電子商務部、行銷部、財務部的「代主管」歸類為「行政幕僚」
    staff_departments = ['人資部', '資訊部', '電子商務部', '行銷部', '財務部']
    if 'job_title_original' in result.columns and 'department' in result.columns:
        mask = (result['job_title_original'] == '代主管') & (result['department'].isin(staff_departments))
        result.loc[mask, 'job_category'] = '行政幕僚'
        result.loc[mask, 'position'] = '行政幕僚'
        adjusted_count = mask.sum()
        if adjusted_count > 0:
            print(f"  ✅ 已調整 {adjusted_count} 位幕僚部門代主管為「行政幕僚」")
    
    # 處理日期欄位（注意：9999 年超出 pandas timestamp 範圍，需特別處理）
    # 先標記 9999 年的記錄（表示在職）
    if 'leave_date' in result.columns:
        def extract_year(x):
            """取得年份，支援 datetime 和字串"""
            if pd.isna(x):
                return None
            if hasattr(x, 'year'):
                return x.year
            try:
                return pd.to_datetime(x).year
            except:
                return None
        
        result['_leave_year'] = result['leave_date'].apply(extract_year)
        result['is_active_by_leave_date'] = result['_leave_year'] >= 9999
        
        # 處理 leave_date：9999 年轉為 None
        def convert_leave_date(x):
            if pd.isna(x):
                return None
            if hasattr(x, 'year') and x.year >= 9999:
                return None
            if hasattr(x, 'date'):
                return x.date()
            try:
                dt = pd.to_datetime(x)
                if pd.notna(dt):
                    return dt.date()
            except:
                pass
            return None
        
        result['leave_date'] = result['leave_date'].apply(convert_leave_date)
        result.drop(columns=['_leave_year'], inplace=True)
    else:
        result['is_active_by_leave_date'] = False
    
    # 處理 hire_date
    if 'hire_date' in result.columns:
        def convert_hire_date(x):
            if pd.isna(x):
                return None
            if hasattr(x, 'date'):
                return x.date()
            try:
                dt = pd.to_datetime(x)
                if pd.notna(dt):
                    return dt.date()
            except:
                pass
            return None
        
        result['hire_date'] = result['hire_date'].apply(convert_hire_date)
    
    # 取得排除的職務類別和工讀生類別
    excluded_categories = jc_config.get('excluded_categories', ['無'])
    part_time_category = jc_config.get('part_time_category', '工讀生')
    
    # 處理員工狀態 - 判斷是否在職
    active_statuses = status_config.get('active_statuses', ['正式員工', '試用員工', '工讀生', '約聘'])
    inactive_statuses = status_config.get('inactive_statuses', ['退休員工', '離職員工', '留職停薪'])
    
    if 'status' in result.columns:
        # 保留原始員工狀態
        result['employee_status_original'] = result['status'].copy()
        
        # 判斷是否為在職（以最後工作日為主要依據）
        def determine_active_status(row):
            job_category = str(row.get('job_category', '')).strip() if pd.notna(row.get('job_category')) else ''
            job_title = str(row.get('job_title_original', '')).strip() if pd.notna(row.get('job_title_original')) else ''
            dept_original = str(row.get('department_original', '')).strip() if pd.notna(row.get('department_original')) else ''
            is_active_by_leave = row.get('is_active_by_leave_date', False)
            leave_date = row.get('leave_date')
            
            # 排除的職務類別（虛擬、雇主）
            if job_category in excluded_categories:
                return '排除'
            
            # 排除「負責人」或空白部門
            if '負責人' in dept_original or dept_original == '':
                return '排除'
            
            # 排除「再生紡織設計部」
            if '再生紡織' in dept_original:
                return '排除'
            
            # 以最後工作日 = 9999/12/31 作為在職判斷依據
            if is_active_by_leave:
                return '在職'
            
            # 如果有離職日且已過，則為離職
            if pd.notna(leave_date):
                if isinstance(leave_date, date) and leave_date <= date.today():
                    return '離職'
            
            return '離職'  # 沒有 9999 且沒有未來離職日，視為離職
        
        result['status'] = result.apply(determine_active_status, axis=1)
        
        # 標記工讀生（根據職務類別）
        if 'job_category' in result.columns:
            result['is_part_time'] = result['job_category'] == part_time_category
        else:
            result['is_part_time'] = False
    else:
        result['status'] = '在職'
        # 檢查是否有 job_category 欄位
        if 'job_category' in result.columns:
            # 排除的職務類別標記為排除
            result['status'] = result['job_category'].apply(
                lambda x: '排除' if x in excluded_categories else '在職'
            )
            result['is_part_time'] = result['job_category'] == part_time_category
        else:
            result['is_part_time'] = False
    
    # 標記移工（根據所屬職務 job_title_original）
    if 'job_title_original' in result.columns:
        result['is_migrant_worker'] = result['job_title_original'] == '移工'
    else:
        result['is_migrant_worker'] = False
    
    # 處理直間接（清理資料）
    if 'labor_type' in result.columns:
        result['labor_type'] = result['labor_type'].apply(
            lambda x: '直接' if '直' in str(x) else ('間接' if '間' in str(x) else '直接') if pd.notna(x) else '直接'
        )
    else:
        result['labor_type'] = '直接'
    
    # 特殊處理1：蔡雅雯（法務及環安課）特殊掛「再生處理部」，不掛人資部
    if 'name' in result.columns:
        special_mask = result['name'] == '蔡雅雯'
        if special_mask.sum() > 0:
            result.loc[special_mask, 'department'] = '再生處理部'
            print(f"  ✅ 已將蔡雅雯特殊處理，改為「再生處理部」")
    
    # 特殊處理1.5：蘇子堯、柯秉杰為技術士（外場人員），非行政幕僚
    if 'name' in result.columns:
        tech_names = ['蘇子堯', '柯秉杰']
        tech_mask = result['name'].isin(tech_names)
        if tech_mask.sum() > 0:
            result.loc[tech_mask, 'job_category'] = '外場人員'
            result.loc[tech_mask, 'position'] = '外場人員'
            print(f"  ✅ 已將蘇子堯、柯秉杰調整為「外場人員」（技術士）")
    
    # 特殊處理2：管理課員工，若在 2025/12/31 前離職，將 department 改為「人資部」
    # 這樣基線計算和離職統計都會正確計入人資部
    cutoff_date = date(2025, 12, 31)
    result['hr_leave_merge'] = False  # 預設不併入（保留此欄位供參考）
    if 'department_original' in result.columns and 'leave_date' in result.columns:
        def should_merge_to_hr(row):
            dept_original = str(row.get('department_original', '')).strip() if pd.notna(row.get('department_original')) else ''
            leave_date = row.get('leave_date')
            
            # 檢查是否為「管理課」（精確匹配，不匹配「電商部行政管理課」等其他部門）
            if dept_original != '管理課':
                return False
            
            # 檢查是否在 2025/12/31 前離職
            if pd.notna(leave_date) and isinstance(leave_date, date):
                return leave_date <= cutoff_date
            
            return False
        
        result['hr_leave_merge'] = result.apply(should_merge_to_hr, axis=1)
        merged_count = result['hr_leave_merge'].sum()
        if merged_count > 0:
            # 將管理課離職員工的 department 改為「人資部」
            result.loc[result['hr_leave_merge'], 'department'] = '人資部'
            print(f"  ✅ 已將 {merged_count} 位管理課離職員工改為「人資部」")
    
    print(f"  ✅ 欄位對應完成，共 {len(result)} 筆資料")
    
    # 計算在職人數（排除工讀生和排除類別）
    active_count = len(result[result['status'] == '在職'])
    part_time_count = len(result[(result['status'] == '在職') & (result['is_part_time'] == True)])
    print(f"  📊 在職人數：{active_count} (含工讀生 {part_time_count})")
    
    return result


def get_external_file_columns(external_path: str = "./數據資料夾") -> List[str]:
    """
    取得外部資料檔案的欄位名稱（供設定欄位對應）
    
    Args:
        external_path: 外部資料夾路徑
        
    Returns:
        欄位名稱列表
    """
    if not os.path.exists(external_path):
        return []
    
    xlsx_files = glob.glob(os.path.join(external_path, "*.xlsx"))
    xls_files = glob.glob(os.path.join(external_path, "*.xls"))
    all_files = xlsx_files + [f for f in xls_files if not f.endswith('.xlsx')]
    
    if not all_files:
        return []
    
    latest_file = max(all_files, key=os.path.getmtime)
    
    try:
        if latest_file.endswith('.xls') and not latest_file.endswith('.xlsx'):
            df = pd.read_excel(latest_file, sheet_name=0, engine='xlrd', nrows=0)
        else:
            df = pd.read_excel(latest_file, sheet_name=0, nrows=0)
        return list(df.columns)
    except Exception:
        return []


def load_departments(file_path: str = "data/departments.xlsx") -> pd.DataFrame:
    """
    載入部門對照表
    
    Args:
        file_path: Excel 檔案路徑
        
    Returns:
        部門對照表 DataFrame
    """
    try:
        df = pd.read_excel(file_path, sheet_name=0)
        return df
        
    except FileNotFoundError:
        print(f"⚠️ 找不到部門檔案：{file_path}，使用預設部門")
        return get_default_departments()
    except Exception as e:
        print(f"⚠️ 讀取部門檔案錯誤：{e}，使用預設部門")
        return get_default_departments()


def get_default_departments() -> pd.DataFrame:
    """取得預設部門對照表"""
    return pd.DataFrame({
        'code': ['GM', 'BD', 'OPS', 'REC', 'INT', 'EC', 'MKT', 'RD', 'FIN', 'HR', 'IT'],
        'name': ['總經理室', '業務開發部', '營運管理部', '再生處理部', '國際事業部', 
                 '電子商務部', '行銷部', '研發部', '財務部', '人資部', '資訊部']
    })


def load_snapshot(file_path: str) -> Optional[Dict[str, Any]]:
    """
    載入歷史快照
    
    Args:
        file_path: 快照檔案路徑
        
    Returns:
        快照資料字典，或 None
    """
    try:
        df = pd.read_excel(file_path, sheet_name=0)
        if len(df) > 0:
            return df.iloc[0].to_dict()
        return None
    except Exception:
        return None


def save_snapshot(data: Dict[str, Any], file_path: str) -> bool:
    """
    儲存快照資料
    
    Args:
        data: 快照資料字典
        file_path: 輸出檔案路徑
        
    Returns:
        是否成功
    """
    try:
        # 確保目錄存在
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        df = pd.DataFrame([data])
        df.to_excel(file_path, index=False, sheet_name='快照')
        return True
    except Exception as e:
        print(f"⚠️ 儲存快照錯誤：{e}")
        return False


def get_snapshot_files(snapshot_path: str = "data/snapshots") -> list:
    """
    取得所有快照檔案列表
    
    Args:
        snapshot_path: 快照目錄路徑
        
    Returns:
        快照檔案路徑列表（依日期排序）
    """
    try:
        if not os.path.exists(snapshot_path):
            return []
            
        files = [
            os.path.join(snapshot_path, f) 
            for f in os.listdir(snapshot_path) 
            if f.endswith('.xlsx')
        ]
        return sorted(files)
    except Exception:
        return []


def get_department_name(code: str, departments_df: pd.DataFrame) -> str:
    """
    取得部門名稱
    
    Args:
        code: 部門代碼
        departments_df: 部門對照表
        
    Returns:
        部門名稱
    """
    match = departments_df[departments_df['code'] == code]
    if len(match) > 0:
        return match.iloc[0]['name']
    return code


def merge_employee_department(
    employees_df: pd.DataFrame, 
    departments_df: pd.DataFrame
) -> pd.DataFrame:
    """
    合併員工與部門資料
    
    Args:
        employees_df: 員工資料
        departments_df: 部門對照表
        
    Returns:
        合併後的 DataFrame（含部門名稱）
    """
    merged = employees_df.merge(
        departments_df,
        left_on='department',
        right_on='code',
        how='left',
        suffixes=('', '_dept')
    )
    merged.rename(columns={'name_dept': 'department_name'}, inplace=True)
    
    # 如果沒有匹配到部門名稱，使用代碼
    merged['department_name'] = merged['department_name'].fillna(merged['department'])
    
    return merged


def load_upcoming_departures() -> pd.DataFrame:
    """
    載入即將離職名單
    從 K:/HR-人資/4.薪酬/2.考勤保險/1.每日出勤/每日出勤總表.xlsx 的「離職」工作表讀取
    
    Returns:
        即將離職人員 DataFrame（事業部、姓名、職稱、離職日）
    """
    file_path = r'K:\HR-人資\4.薪酬\2.考勤保險\1.每日出勤\每日出勤總表.xlsx'
    
    try:
        df = pd.read_excel(file_path, sheet_name='離職')
        
        # 轉換離職日為日期（錯誤的設為NaT）
        df['離職日'] = pd.to_datetime(df['離職日'], errors='coerce').dt.date
        
        # 排除無效日期和缺少部門/姓名的資料
        df = df[df['離職日'].notna()]
        df = df[df['部門'].notna()]
        df = df[df['姓名'].notna()]
        
        # 篩選即將離職（離職日在今天或之後）
        today = date.today()
        upcoming = df[df['離職日'] >= today].copy()
        
        # 載入事業部對應表
        config = load_config()
        external_path = config.get('external_data', {}).get('path', './數據資料夾')
        bu_mapping = load_business_unit_mapping(external_path, config)
        
        # 將部門對應到事業部
        if bu_mapping:
            upcoming['事業部'] = upcoming['部門'].map(bu_mapping).fillna(upcoming['部門'])
        else:
            upcoming['事業部'] = upcoming['部門']
        
        # 只保留需要的欄位並排序
        result = upcoming[['事業部', '姓名', '職稱', '離職日']].sort_values('離職日')
        
        return result
        
    except FileNotFoundError:
        print(f"⚠️ 找不到即將離職名單檔案：{file_path}")
        return pd.DataFrame(columns=['事業部', '姓名', '職稱', '離職日'])
    except Exception as e:
        print(f"⚠️ 讀取即將離職名單錯誤：{e}")
        return pd.DataFrame(columns=['事業部', '姓名', '職稱', '離職日'])
