@echo off
chcp 65001 >nul
title PyInstaller Packager

echo Installing PyInstaller...
pip install pyinstaller -q

echo.
echo Packaging... (takes 2-5 minutes)
echo.

pyinstaller ^
  --name "凭证生成工具" ^
  --onedir ^
  --noconsole ^
  --add-data "app.py;." ^
  --add-data "company_manager.py;." ^
  --add-data "processor;processor" ^
  --add-data "utils;utils" ^
  --add-data "requirements.txt;." ^
  --collect-all streamlit ^
  --collect-all pandas ^
  --collect-all openpyxl ^
  --hidden-import streamlit ^
  --hidden-import streamlit.web.cli ^
  --hidden-import streamlit.runtime.scriptrunner ^
  launcher.py

echo.
if exist "dist\凭证生成工具\凭证生成工具.exe" (
    echo [OK] Done! EXE is in: dist\凭证生成工具\
    echo Send the entire "凭证生成工具" folder to users.
) else (
    echo [ERROR] Build failed. See output above.
)
pause
