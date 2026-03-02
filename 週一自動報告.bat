@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ================================================== 
echo 大豐環保人力報告 - 每週一自動執行
echo 執行時間: %date% %time%
echo ==================================================

"C:\Users\chiehyi\AppData\Local\Programs\Python\Python313\python.exe" auto_report.py

if %errorlevel% equ 0 (
    echo.
    echo ✅ 報告執行完成！
) else (
    echo.
    echo ❌ 報告執行失敗，錯誤碼: %errorlevel%
)

echo.
echo 按任意鍵關閉...
pause >nul
