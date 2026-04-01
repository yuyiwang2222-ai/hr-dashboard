@echo off
chcp 65001 >nul
cd /d "%~dp0"
"C:\Users\chiehyi\AppData\Local\Programs\Python\Python313\python.exe" scripts\run_scheduled_monday.py --phase email