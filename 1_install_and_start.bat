@echo off
title Voucher Tool Setup
cd /d "%~dp0"

if exist python\python.exe goto already_installed

echo ================================================
echo  First time setup - downloading Python runtime
echo  This takes 2-3 minutes, please wait...
echo ================================================
echo.

:: Download Python embeddable
echo [1/4] Downloading Python...
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip' -OutFile 'python_embed.zip'"
if errorlevel 1 goto dl_error

echo [2/4] Extracting Python...
powershell -Command "Expand-Archive -Path 'python_embed.zip' -DestinationPath 'python' -Force"
del python_embed.zip

:: Enable pip by editing pth file
echo [3/4] Setting up pip...
echo import site >> python\python311._pth
powershell -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile 'get-pip.py'"
python\python.exe get-pip.py --quiet
del get-pip.py

:: Install dependencies
echo [4/4] Installing dependencies (streamlit, pandas, openpyxl)...
python\python.exe -m pip install streamlit pandas openpyxl --quiet --no-warn-script-location

echo.
echo [OK] Setup complete!
echo ================================================

:already_installed
echo Starting Voucher Tool...
start /b cmd /c "timeout /t 3 /nobreak >nul && start http://localhost:8501"
python\python.exe -m streamlit run app.py --server.headless true --browser.gatherUsageStats false
goto end

:dl_error
echo.
echo [ERROR] Download failed. Please check your internet connection.
pause
exit /b 1

:end
pause
