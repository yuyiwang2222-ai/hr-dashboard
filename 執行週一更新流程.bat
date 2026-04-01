@echo off
chcp 65001 >nul
cd /d "%~dp0"
"C:\Users\chiehyi\AppData\Local\Programs\Python\Python313\python.exe" scripts\monday_workflow.py --force --once
pause
