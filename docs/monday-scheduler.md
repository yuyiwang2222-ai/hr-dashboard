# 自動排程（每週一 + 月初）

## 排程時間

| 觸發時機 | Prepare 階段 | Email 階段 |
|---------|-------------|-----------|
| 每週一 | 09:00 | 09:30 |
| 每月 1 日 | 09:00 | 09:30 |

> 若 1 日剛好是週一，只會執行一次（不會重複）。

## 註冊排程

在 PowerShell 執行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register_auto_tasks.ps1
```

## 移除排程

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\unregister_auto_tasks.ps1
```

## 排程內容

### Prepare 階段（09:00）
1. 從 `J:\HR-人資\4.薪酬\2.考勤保險\1.每日出勤\` 複製每日出勤總表
2. 從 HRM 系統匯出員工人數
3. 同步資料到 `data/` 資料夾
4. 執行每週任務（產生快照）
5. 產生自動報告

### Email 階段（09:30）
1. 預設不自動寄送（避免未確認內容就直接發信）
2. 由你在儀表板手動勾選確認後按「傳送」
3. 若需要命令列強制寄送，可執行：
	`python scripts/run_scheduled_monday.py --phase email --confirm-send`

## 注意事項

- `09:00` 的 HRM 匯出會操作桌面畫面，建議該時段不要手動使用同一台電腦
- 電腦需保持登入狀態
- `J:` 網路磁碟需可存取
- HRM 程式需可正常使用

## 環境變數設定

在 `.env` 檔案中設定：

```bash
# 每日出勤總表來源路徑
ATTENDANCE_SOURCE_PATH=J:\HR-人資\4.薪酬\2.考勤保險\1.每日出勤\每日出勤總表.xlsx

# HRM 自動化設定
HRM_LAUNCHER_PATH=C:\Program Files (x86)\Digiwin\DigiwinHR\Launcher.exe
HRM_USERNAME=你的帳號
HRM_PASSWORD=你的密碼
HRM_COMPANY=公司名稱
```
