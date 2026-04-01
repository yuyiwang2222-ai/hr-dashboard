# 每週一 09:00 / 09:30 自動排程

在 PowerShell 執行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register_monday_tasks.ps1
```

移除排程：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\unregister_monday_tasks.ps1
```

排程內容：
- `HRDashboard-Monday-Prepare`：每週一 09:00 執行 `--phase prepare`
- `HRDashboard-Monday-Email`：每週一 09:30 執行 `--phase email`

注意：
- `09:00` 的 HRM 匯出會操作桌面畫面，建議該時段不要手動使用同一台電腦。
- 電腦需保持登入狀態，且 `J:` 網路磁碟與 HRM 程式可正常使用。
