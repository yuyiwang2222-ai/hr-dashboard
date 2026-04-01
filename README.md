# 大豐環保人力變化視覺化儀表板

即時追蹤公司人力變化的互動式儀表板，支援部門分析、異常警報、自動化週報。

## 🚀 快速開始

### 1. 安裝套件

```bash
pip install -r requirements.txt
```

### 2. 初始化測試資料

```bash
python setup_data.py
```

### 3. 啟動儀表板

```bash
streamlit run app.py
```

或直接執行：

```bash
run.bat
```

瀏覽器將自動開啟 `http://localhost:8501`

---

## 📁 專案結構

```
Agents-人力變化/
├── app.py                 # Streamlit 主程式
├── requirements.txt       # 套件清單
├── config.yaml           # 設定檔
├── setup_data.py         # 測試資料產生器
├── run.bat               # 一鍵啟動
├── README.md             # 本文件
├── PLAN.md               # 實作計劃
├── SPEC.md               # 規格書
├── data/
│   ├── employees.xlsx    # 員工主檔
│   ├── departments.xlsx  # 部門對照表
│   └── snapshots/        # 歷史快照
├── src/
│   ├── __init__.py
│   ├── data_loader.py    # 資料讀取模組
│   ├── metrics.py        # 指標計算模組
│   ├── charts.py         # 圖表產生模組
│   ├── alerts.py         # 警報偵測模組
│   └── email_report.py   # Email 週報模組
└── scripts/
    └── weekly_job.py     # 每週排程任務
```

---

## 📊 功能說明

### 首頁總覽
- **KPI 卡片**：總人數、本週入職、本週離職、月離職率
- **人數趨勢圖**：過去 12 週人數變化
- **部門人數圖**：各部門人數（可切換顯示直接/間接）
- **人員分布圖**：職務分布、僱用類型、直間接比例

### 部門分析
- 選擇部門查看詳細統計
- 部門職務組成
- 部門直間接比例
- 部門人員名單

### 異動名單
- 本週/本月/自訂期間的入職離職名單
- 可查看詳細人員資訊

### 警報看板
- 即時異常警報顯示
- 離職率過高警報
- 部門離職異常警報
- 人力大幅變動警報

---

## ⚙️ 設定說明

編輯 `config.yaml` 調整系統設定：

### 警報閾值

```yaml
alerts:
  monthly_turnover_rate:
    warning: 5      # 月離職率 > 5% 觸發黃燈
    critical: 10    # 月離職率 > 10% 觸發紅燈
  department_monthly_leaves:
    warning: 3      # 單部門月離職 > 3 人觸發警報
```

### Email 設定

```yaml
email:
  enabled: true
  smtp_server: "smtp.company.com"
  smtp_port: 587
  sender: "hr-system@company.com"
  recipients:
    - "manager@company.com"
```

> ⚠️ Email 密碼請設定環境變數 `EMAIL_PASSWORD`

---

## ⏰ 自動化排程

### 設定 Windows Task Scheduler

1. 開啟「工作排程器」
2. 建立基本工作
3. 設定：
   - **名稱**：人力儀表板週報
   - **觸發條件**：每週一 09:00
   - **動作**：啟動程式
   - **程式**：`python`
   - **引數**：`scripts/weekly_job.py`
   - **起始位置**：`{專案目錄完整路徑}`

### 手動執行週報

```bash
python scripts/weekly_job.py
```

---

## 📝 資料維護

### 員工資料格式（employees.xlsx）

| 欄位 | 說明 | 必填 | 範例 |
|------|------|------|------|
| employee_id | 員工編號 | ✓ | TF0001 |
| name | 姓名 | ✓ | 王小明 |
| department | 部門代碼 | ✓ | IT |
| position | 職務 | ✓ | 工程師 |
| employment_type | 僱用類型 | ✓ | 正職/約聘/實習 |
| labor_type | 直間接 | ✓ | 直接/間接 |
| hire_date | 到職日 | ✓ | 2024-01-15 |
| leave_date | 離職日 |  | 2024-12-31（空白=在職）|
| status | 狀態 | ✓ | 在職/離職 |

### 部門代碼對照

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

---

## 🔧 故障排除

### 問題：找不到員工資料
```
⚠️ 找不到員工資料！
```
**解決**：執行 `python setup_data.py` 建立測試資料

### 問題：套件安裝失敗
**解決**：確認 Python 版本 >= 3.9，執行 `pip install --upgrade pip`

### 問題：圖表無法顯示
**解決**：重新安裝 plotly：`pip install --upgrade plotly`

---

## 📞 聯絡資訊

如有問題請聯繫資訊部。

---

© 2026 大豐環保科技股份有限公司

## 🔁 VS Code 週一自動流程

現在這個專案支援「每週一更新兩個 Excel 後，自動同步資料並開本機網頁」。

### 自動執行方式

1. 用 VS Code 開啟整個 `Agents-人力變化` 專案資料夾。
2. 第一次開啟如果看到 `Allow Automatic Tasks`，請選 `Allow`。
3. 把以下兩個檔案覆蓋成最新版本：
   - `數據資料夾/員工人數.xlsx`
   - `數據資料夾/每日出勤總表.xlsx`
4. VS Code 背景 task 會自動：
   - 執行 `sync_data.py` 同步到 `data/`
   - 執行 `scripts/weekly_job.py`
   - 啟動或重啟本機 Streamlit
   - 自動開啟 `http://127.0.0.1:8501`

### 手動補跑

如果你不想等監看自動偵測，也可以直接執行：

```bash
python scripts/monday_workflow.py --force --once
```

或直接雙擊：

- `執行週一更新流程.bat`

### 說明

- 目前預設是更新本機網頁，不會自動 `git commit` / `git push`。
- 你提供的雲端網址如果要跟著每週自動更新，可以再另外接一個「發佈到 GitHub / Streamlit Cloud」流程。
## ⏰ 週一 09:00 / 09:30 自動流程

目前專案已支援把 HRM 匯出與週報寄送拆成兩個排程：

### 09:00 更新資料

執行：`run-0900-update.bat`

這一步會自動：
- 從 `J:\HR-人資\4.薪酬\2.考勤保險\1.每日出勤\每日出勤總表.xlsx` 複製最新出勤檔
- 開啟 Digiwin HRM 並匯出 `員工人數.xlsx`
- 同步到 `data/`
- 執行 `weekly_job.py`
- 產生最新 PDF 與更新雲端資料，但先不寄信

### 09:30 寄送週報

執行：`run-0930-email.bat`

這一步會直接寄出當天最新 PDF 報告。

### 建議的 Windows Task Scheduler 設定

建立兩個工作：

1. `人力週報-09點更新資料`
   - 時間：每週一 09:00
   - 程式：`run-0900-update.bat`

2. `人力週報-09點30寄信`
   - 時間：每週一 09:30
   - 程式：`run-0930-email.bat`

### HRM 自動化前置設定

請先在 `.env` 設定：

- `HRM_USERNAME`
- `HRM_PASSWORD`
- `HRM_COMPANY`（若登入需要）
- `ATTENDANCE_SOURCE_PATH`

若 HRM 視窗版面和目前截圖不一致，`scripts/export_hrm_employees.py` 裡的座標常數可能需要微調。