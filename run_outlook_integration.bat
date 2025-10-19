@echo off
chcp 65001 >nul
echo ========================================
echo 🤖 AI Email Manager - Outlook Integration
echo ========================================
echo.
echo מפעיל את האינטגרציה עם Outlook...
echo.

cd /d "%~dp0"

REM בדיקה אם venv קיים
if not exist "venv\Scripts\python.exe" (
    echo ❌ סביבה וירטואלית לא נמצאה
    echo הפעל תחילה: install.bat
    pause
    exit /b 1
)

REM הפעלת האינטגרציה
venv\Scripts\python.exe outlook_integration.py

pause















