@echo off
cd /d "%~dp0"
title Voucher Tool
start /b cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:8501"
python\python.exe -m streamlit run app.py --server.headless true --browser.gatherUsageStats false
pause
