@echo off
chcp 65001 >nul
echo ============================================
echo  大豐環保人力儀表板
echo ============================================
echo.

REM 檢查資料檔案是否存在
if not exist "data\employees.xlsx" (
    echo [!] 找不到員工資料，正在初始化測試資料...
    python setup_data.py
    echo.
)

echo [*] 啟動儀表板...
echo.
echo     電腦瀏覽：http://localhost:8501
echo     手機瀏覽：請查看下方 Network URL
echo.
echo [!] 手機需與電腦在同一個 Wi-Fi 網路
echo [*] 按 Ctrl+C 停止服務
echo.

streamlit run app.py --server.address 0.0.0.0 --server.port 8501

pause
