# 大豐環保人力變化視覺化儀表板 - 實作計劃

> 建立日期：2026-02-12  
> 狀態：實作中

## 📋 專案概述

使用 Python + Streamlit 建立互動式人力儀表板，追蹤 11 個部門的人員進出、分布與趨勢。資料存於 Excel（便於人資維護），搭配每週自動快照、離職率警報、Email 週報功能。

### 目標
- 即時呈現人力變化數據
- 支援部門、職務、直/間接人員多維度分析
- 異常警報（離職率過高）
- 每週自動更新 + Email 週報

---

## 📁 專案結構

```
Agents-人力變化/
├── app.py                 # Streamlit 主程式
├── requirements.txt       # 套件清單
├── run.bat               # 一鍵啟動
├── config.yaml           # 設定檔（警報閾值、Email設定）
├── README.md             # 使用說明
├── PLAN.md               # 本文件
├── SPEC.md               # 規格書
├── data/
│   ├── employees.xlsx    # 員工主檔
│   ├── departments.xlsx  # 部門對照表
│   └── snapshots/        # 歷史快照
├── src/
│   ├── __init__.py
│   ├── data_loader.py    # 資料讀取
│   ├── metrics.py        # 指標計算
│   ├── charts.py         # 圖表產生
│   ├── alerts.py         # 異常檢測
│   └── email_report.py   # Email 寄送
└── scripts/
    └── weekly_job.py     # 排程任務
```

---

## 📊 資料結構設計

### employees.xlsx（員工主檔）

| 欄位 | 說明 | 範例 |
|------|------|------|
| employee_id | 員工編號 | TF0001 |
| name | 姓名 | 王小明 |
| department | 部門代碼 | OPS |
| position | 職務 | 廠長 / 課長 / 工程師 / 作業員 |
| employment_type | 僱用類型 | 正職 / 約聘 / 實習 |
| labor_type | 直間接 | 直接 / 間接 |
| hire_date | 到職日 | 2024-01-15 |
| leave_date | 離職日 | (空白=在職) |
| status | 狀態 | 在職 / 離職 |

### departments.xlsx（部門對照表）

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

## 🖥️ 儀表板功能設計

### 首頁 Dashboard
- **KPI 卡片**：總人數、本週入職、本週離職、月離職率
- **人數趨勢圖**：過去 12 週折線圖
- **部門人數圖**：長條圖（可篩選直/間接）
- **職務分布圖**：圓餅圖
- **僱用類型分布**：圓餅圖

### 部門分析頁
- 各部門人數熱力圖
- 部門進出趨勢
- 部門直間接比例

### 異動名單頁
- 本週/本月入職清單
- 本週/本月離職清單
- 可匯出 Excel

### 警報看板
- 超標警報顯示（紅燈/黃燈）
- 警報歷史記錄

---

## ⚠️ 異常警報機制

### 警報閾值（config.yaml）
```yaml
alerts:
  monthly_turnover_rate:
    warning: 5    # 黃燈：月離職率 > 5%
    critical: 10  # 紅燈：月離職率 > 10%
  department_leaves:
    warning: 3    # 單部門月離職 > 3 人
```

---

## 📧 Email 週報功能

### 週報內容
- 本週人力摘要（總人數、進出人數）
- 關鍵指標變化
- 異常警報提醒
- 附件：儀表板截圖

### 設定（config.yaml）
```yaml
email:
  smtp_server: smtp.company.com
  smtp_port: 587
  sender: hr-system@company.com
  recipients:
    - manager@company.com
    - hr@company.com
```

---

## ⏰ 自動化排程

### 每週任務（週一 09:00）
1. 產生本週人力快照 → `data/snapshots/YYYY-WW.xlsx`
2. 執行異常檢測
3. 寄送 Email 週報

### Windows Task Scheduler 設定
- 程式：`python`
- 參數：`scripts/weekly_job.py`
- 起始位置：專案根目錄

---

## 🔧 技術選型

| 項目 | 選擇 | 原因 |
|------|------|------|
| 框架 | Streamlit | 快速開發、互動性佳 |
| 資料源 | Excel | 人資熟悉、易維護 |
| 圖表 | Plotly | 互動式、美觀 |
| 排程 | Windows Task Scheduler | 本機部署、簡單可靠 |

---

## ✅ 實作步驟

- [x] 1. 建立專案目錄結構
- [x] 2. 建立 PLAN.md 和 SPEC.md
- [ ] 3. 建立 requirements.txt 並安裝套件
- [ ] 4. 設計並建立 Excel 資料範本
- [ ] 5. 開發 src/data_loader.py
- [ ] 6. 開發 src/metrics.py
- [ ] 7. 開發 src/charts.py
- [ ] 8. 開發主程式 app.py
- [ ] 9. 開發 src/alerts.py
- [ ] 10. 開發 src/email_report.py
- [ ] 11. 開發 scripts/weekly_job.py
- [ ] 12. 建立 config.yaml
- [ ] 13. 撰寫 README.md
- [ ] 14. 建立 run.bat
- [ ] 15. 測試與驗證

---

## 🧪 驗證方式

1. **資料驗證**：匯入測試員工資料，確認讀取正確
2. **圖表驗證**：啟動 Streamlit，檢查所有圖表互動功能
3. **警報驗證**：模擬高離職率情境，確認警報觸發
4. **Email 驗證**：手動執行週報腳本，確認郵件寄送
5. **排程驗證**：設定 Task Scheduler，確認自動執行

---

## 📝 決策紀錄

| 決策 | 選擇 | 原因 |
|------|------|------|
| 框架選擇 | Streamlit > Dash | 更快開發、語法簡潔 |
| 資料儲存 | Excel > 資料庫 | 人資熟悉、易維護 |
| 部署方式 | 本機 > 雲端 | 先驗證功能、降低複雜度 |
| 職級欄位 | 移除 | 用戶確認不需要 |
| 部門分類 | 移除 | 用戶確認不需要 |

---

## 🚀 後續擴展

- 串接公司人資系統 API（自動同步）
- 部署至內網伺服器（多人存取）
- 新增招募進度追蹤
- 升級至 SQLite/PostgreSQL 資料庫
