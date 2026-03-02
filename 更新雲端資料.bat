@echo off
chcp 65001 >nul
echo ========================================
echo 大豐環保人力儀表板 - 更新並推送到雲端
echo ========================================
echo.

REM 步驟 1：同步資料
echo [步驟 1/3] 同步資料檔案...
python sync_data.py
echo.

REM 步驟 2：Git 提交
echo [步驟 2/3] 提交變更到 Git...
git add data/
git commit -m "更新員工資料 %date%"
echo.

REM 步驟 3：推送到 GitHub
echo [步驟 3/3] 推送到 GitHub（Streamlit Cloud 會自動更新）...
git push
echo.

echo ========================================
echo ✅ 完成！Streamlit Cloud 將在約 1-2 分鐘內自動更新
echo ========================================
pause
