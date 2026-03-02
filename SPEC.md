# 大豐環保人力變化視覺化系統 - 規格書

> 版本：1.0  
> 建立日期：2026-02-12  
> 文件類型：系統規格書

---

## 1. 系統概述

### 1.1 目的
建立一套人力變化視覺化系統，協助管理層即時掌握組織人力動態，支援決策分析。

### 1.2 範圍
- 涵蓋公司 11 個部門、200-500 名員工
- 追蹤人員進出、部門分布、職務分布、直間接人員比例
- 提供異常警報與自動化週報

### 1.3 使用者
| 角色 | 使用情境 |
|------|----------|
| 人資部 | 維護員工資料、查看報表 |
| 部門主管 | 查看部門人力狀況 |
| 高階主管 | 查看整體人力趨勢、接收週報 |

---

## 2. 功能規格

### 2.1 儀表板功能

#### 2.1.1 首頁總覽
| 功能 | 說明 | 優先級 |
|------|------|--------|
| KPI 卡片 | 顯示總人數、本週入職、本週離職、月離職率 | P0 |
| 人數趨勢圖 | 過去 12 週人數變化折線圖 | P0 |
| 部門人數圖 | 各部門人數長條圖，可篩選直/間接 | P0 |
| 職務分布圖 | 各職務人數圓餅圖 | P1 |
| 僱用類型分布 | 正職/約聘/實習圓餅圖 | P1 |

#### 2.1.2 部門分析
| 功能 | 說明 | 優先級 |
|------|------|--------|
| 部門篩選器 | 下拉選單選擇部門 | P0 |
| 部門人數趨勢 | 單一部門歷史趨勢 | P1 |
| 部門直間接比例 | 圓餅圖呈現 | P1 |
| 部門職務組成 | 該部門各職務人數 | P1 |

#### 2.1.3 異動名單
| 功能 | 說明 | 優先級 |
|------|------|--------|
| 期間篩選 | 本週/本月/自訂區間 | P0 |
| 入職名單 | 表格顯示新進人員 | P0 |
| 離職名單 | 表格顯示離職人員 | P0 |
| 匯出功能 | 匯出為 Excel | P2 |

#### 2.1.4 警報看板
| 功能 | 說明 | 優先級 |
|------|------|--------|
| 即時警報 | 顯示目前觸發的警報 | P0 |
| 警報歷史 | 過去警報記錄 | P2 |
| 警報詳情 | 點擊查看詳細數據 | P2 |

### 2.2 異常警報功能

#### 2.2.1 警報類型
| 警報名稱 | 觸發條件 | 等級 |
|----------|----------|------|
| 高離職率警告 | 月離職率 > 5% | 黃燈 |
| 高離職率危急 | 月離職率 > 10% | 紅燈 |
| 部門異常離職 | 單部門月離職 > 3 人 | 黃燈 |
| 人力大幅變動 | 週人數變化 > 5% | 黃燈 |

#### 2.2.2 警報行為
- 儀表板即時顯示
- Email 通知（可設定收件人）
- 記錄至警報歷史

### 2.3 自動化功能

#### 2.3.1 每週快照
| 項目 | 說明 |
|------|------|
| 執行時間 | 每週一 09:00 |
| 輸出檔案 | `data/snapshots/YYYY-WW.xlsx` |
| 內容 | 當週人力統計數據 |

#### 2.3.2 Email 週報
| 項目 | 說明 |
|------|------|
| 執行時間 | 每週一 09:30（快照後） |
| 收件人 | 可於 config.yaml 設定 |
| 內容 | 人力摘要、關鍵指標、警報、附件圖表 |

---

## 3. 資料規格

### 3.1 員工主檔 (employees.xlsx)

| 欄位名稱 | 資料類型 | 必填 | 說明 | 驗證規則 |
|----------|----------|------|------|----------|
| employee_id | String | Y | 員工編號 | 唯一值，格式 TFxxxx |
| name | String | Y | 姓名 | 最多 50 字元 |
| department | String | Y | 部門代碼 | 須存在於 departments.xlsx |
| position | String | Y | 職務名稱 | 最多 50 字元 |
| employment_type | String | Y | 僱用類型 | 正職/約聘/實習 |
| labor_type | String | Y | 直間接 | 直接/間接 |
| hire_date | Date | Y | 到職日 | YYYY-MM-DD 格式 |
| leave_date | Date | N | 離職日 | YYYY-MM-DD 格式，空白=在職 |
| status | String | Y | 狀態 | 在職/離職 |

### 3.2 部門對照表 (departments.xlsx)

| 欄位名稱 | 資料類型 | 必填 | 說明 |
|----------|----------|------|------|
| code | String | Y | 部門代碼（唯一） |
| name | String | Y | 部門名稱 |

### 3.3 部門清單

| 代碼 | 部門名稱 |
|------|----------|
| GM | 總經理室 |
| BD | 業務開發部 |
| OPS | 營運管理部 |
| REC | 再生處理部 |
| INT | 國際事業部 |
| EC | 電子商務部 |
| MKT | 行銷部 |
| RD | 研發部 |
| FIN | 財務部 |
| HR | 人資部 |
| IT | 資訊部 |

### 3.4 歷史快照 (snapshots/YYYY-WW.xlsx)

| 欄位名稱 | 資料類型 | 說明 |
|----------|----------|------|
| snapshot_date | Date | 快照日期 |
| total_headcount | Integer | 總人數 |
| by_department | JSON | 各部門人數 |
| by_position | JSON | 各職務人數 |
| by_employment_type | JSON | 各僱用類型人數 |
| by_labor_type | JSON | 直間接人數 |
| weekly_hires | Integer | 本週入職 |
| weekly_leaves | Integer | 本週離職 |
| monthly_turnover_rate | Float | 月離職率 |

---

## 4. 介面規格

### 4.1 頁面結構

```
┌─────────────────────────────────────────────────────┐
│  🏢 大豐環保人力儀表板          [日期篩選] [重整]    │
├─────────────────────────────────────────────────────┤
│  📊 首頁 │ 📈 部門分析 │ 📋 異動名單 │ ⚠️ 警報     │
├─────────────────────────────────────────────────────┤
│                                                     │
│  ┌─────┐ ┌─────┐ ┌─────┐ ┌─────┐                   │
│  │總人數│ │入職 │ │離職 │ │離職率│  ← KPI 卡片     │
│  │ 350 │ │ +5  │ │ -2  │ │3.2% │                   │
│  └─────┘ └─────┘ └─────┘ └─────┘                   │
│                                                     │
│  ┌─────────────────────────────────────────────┐   │
│  │          人數趨勢圖（12週）                  │   │
│  │     📈 ──────────────────────────           │   │
│  └─────────────────────────────────────────────┘   │
│                                                     │
│  ┌──────────────────┐  ┌──────────────────┐        │
│  │   部門人數圖     │  │   職務分布圖     │        │
│  │   📊             │  │   🥧             │        │
│  └──────────────────┘  └──────────────────┘        │
│                                                     │
└─────────────────────────────────────────────────────┘
```

### 4.2 互動元件

| 元件 | 類型 | 功能 |
|------|------|------|
| 日期篩選器 | Date Range Picker | 篩選資料期間 |
| 部門篩選器 | Multi-Select | 篩選部門（可多選） |
| 直間接切換 | Radio Button | 全部/直接/間接 |
| 重整按鈕 | Button | 重新載入資料 |
| 匯出按鈕 | Button | 匯出當前資料 |

### 4.3 圖表規格

| 圖表 | 類型 | 資料維度 | 互動功能 |
|------|------|----------|----------|
| 人數趨勢 | Line Chart | 日期 x 人數 | Hover 顯示數值 |
| 部門人數 | Bar Chart | 部門 x 人數 | 點擊篩選、排序 |
| 職務分布 | Pie Chart | 職務 x 人數 | Hover 顯示百分比 |
| 僱用類型 | Pie Chart | 類型 x 人數 | Hover 顯示百分比 |

### 4.4 色彩規範

| 用途 | 色碼 | 說明 |
|------|------|------|
| 主色 | #1E88E5 | 品牌藍 |
| 成功/入職 | #4CAF50 | 綠色 |
| 警告/黃燈 | #FFC107 | 黃色 |
| 危急/紅燈 | #F44336 | 紅色 |
| 離職 | #9E9E9E | 灰色 |
| 背景 | #FAFAFA | 淺灰 |

---

## 5. 技術規格

### 5.1 系統需求

| 項目 | 規格 |
|------|------|
| 作業系統 | Windows 10/11 |
| Python | 3.9+ |
| 記憶體 | 4GB+ |
| 磁碟空間 | 500MB+ |
| 瀏覽器 | Chrome/Edge（最新版） |

### 5.2 套件相依

```
streamlit>=1.28.0
pandas>=2.0.0
openpyxl>=3.1.0
plotly>=5.18.0
pyyaml>=6.0
python-dateutil>=2.8.0
```

### 5.3 設定檔結構 (config.yaml)

```yaml
# 系統設定
system:
  app_title: "大豐環保人力儀表板"
  data_path: "./data"
  snapshot_path: "./data/snapshots"
  
# 警報設定
alerts:
  monthly_turnover_rate:
    warning: 5
    critical: 10
  department_monthly_leaves:
    warning: 3
  weekly_headcount_change:
    warning: 5

# Email 設定
email:
  enabled: true
  smtp_server: "smtp.company.com"
  smtp_port: 587
  use_tls: true
  sender: "hr-system@company.com"
  sender_password: "${EMAIL_PASSWORD}"  # 使用環境變數
  recipients:
    - "manager@company.com"
    - "hr@company.com"
  
# 排程設定
schedule:
  snapshot_day: "monday"
  snapshot_time: "09:00"
  report_time: "09:30"
```

### 5.4 API 規格（內部模組）

#### data_loader.py
```python
def load_employees(file_path: str) -> pd.DataFrame
def load_departments(file_path: str) -> pd.DataFrame
def load_snapshot(file_path: str) -> dict
def save_snapshot(data: dict, file_path: str) -> None
```

#### metrics.py
```python
def get_current_headcount(df: pd.DataFrame, filters: dict = None) -> int
def get_period_changes(df: pd.DataFrame, start_date: date, end_date: date) -> dict
def get_turnover_rate(df: pd.DataFrame, period: str = "monthly") -> float
def get_department_stats(df: pd.DataFrame) -> pd.DataFrame
def get_position_stats(df: pd.DataFrame) -> pd.DataFrame
def get_labor_type_stats(df: pd.DataFrame) -> pd.DataFrame
```

#### alerts.py
```python
def check_all_alerts(df: pd.DataFrame, config: dict) -> list[Alert]
def get_alert_history(days: int = 30) -> list[Alert]
def format_alert_message(alert: Alert) -> str
```

#### email_report.py
```python
def generate_report_content(metrics: dict, alerts: list) -> str
def send_email(subject: str, body: str, attachments: list = None) -> bool
def send_weekly_report() -> bool
```

---

## 6. 測試規格

### 6.1 單元測試

| 模組 | 測試項目 | 預期結果 |
|------|----------|----------|
| data_loader | 讀取有效 Excel | 回傳 DataFrame |
| data_loader | 讀取無效路徑 | 拋出 FileNotFoundError |
| metrics | 計算人數 | 回傳正確整數 |
| metrics | 計算離職率 | 回傳 0-100 之間 |
| alerts | 超過閾值 | 產生警報物件 |
| alerts | 未超過閾值 | 回傳空列表 |

### 6.2 整合測試

| 測試案例 | 步驟 | 預期結果 |
|----------|------|----------|
| 儀表板啟動 | 執行 streamlit run app.py | 無錯誤開啟瀏覽器 |
| 資料載入 | 放入測試資料 | 圖表正確顯示 |
| 篩選功能 | 選擇部門篩選 | 圖表即時更新 |
| 警報觸發 | 模擬高離職率 | 警報看板顯示 |
| 週報寄送 | 執行 weekly_job.py | 收到測試郵件 |

### 6.3 測試資料

建立 50 筆測試員工資料：
- 涵蓋所有 11 個部門
- 包含各種僱用類型
- 包含直接與間接人員
- 有近期入職/離職記錄

---

## 7. 部署規格

### 7.1 本機部署步驟

1. 安裝 Python 3.9+
2. 執行 `pip install -r requirements.txt`
3. 準備資料檔案至 `data/` 目錄
4. 設定 `config.yaml`
5. 執行 `run.bat` 或 `streamlit run app.py`

### 7.2 排程設定

Windows Task Scheduler 設定：
| 項目 | 值 |
|------|-----|
| 名稱 | 人力儀表板週報 |
| 觸發 | 每週一 09:00 |
| 動作 | 啟動程式 |
| 程式 | python |
| 參數 | scripts/weekly_job.py |
| 起始 | {專案根目錄} |

### 7.3 未來雲端部署選項

| 平台 | 適用情境 | 預估成本 |
|------|----------|----------|
| Streamlit Cloud | 外部展示 | 免費 |
| Azure App Service | 企業內網 | $50/月 |
| 公司內部伺服器 | 完全控制 | 硬體成本 |

---

## 8. 安全規格

### 8.1 資料安全
- 員工資料僅存於本機
- Excel 檔案可設定密碼保護
- 不上傳至雲端（除非明確部署）

### 8.2 存取控制（未來擴展）
- Streamlit 支援基本驗證
- 可整合公司 SSO

### 8.3 敏感資料處理
- Email 密碼使用環境變數
- 不在程式碼中硬編碼憑證

---

## 9. 維護規格

### 9.1 日常維護
| 任務 | 頻率 | 負責人 |
|------|------|--------|
| 更新員工資料 | 即時 | 人資部 |
| 檢查警報 | 每日 | 人資部 |
| 確認週報寄送 | 每週 | IT |

### 9.2 定期維護
| 任務 | 頻率 | 說明 |
|------|------|------|
| 清理舊快照 | 每季 | 保留最近 52 週 |
| 更新套件 | 每月 | pip upgrade |
| 備份資料 | 每週 | 複製 data/ 目錄 |

---

## 10. 附錄

### 10.1 名詞定義

| 名詞 | 定義 |
|------|------|
| 在職人數 | status = "在職" 的員工數 |
| 離職率 | 期間離職人數 / 期初在職人數 × 100% |
| 直接人員 | 直接參與生產/服務的員工 |
| 間接人員 | 支援性質的員工（行政、管理） |

### 10.2 版本歷史

| 版本 | 日期 | 變更說明 |
|------|------|----------|
| 1.0 | 2026-02-12 | 初版建立 |

---

*文件結束*
