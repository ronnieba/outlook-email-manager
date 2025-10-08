@echo off
chcp 65001 >nul
echo.
echo ========================================
echo    AI Email Manager - Simple Installation
echo ========================================
echo.

echo Step 1: Checking requirements...
echo.

:: Check Python
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ Python not installed
    echo Please install Python 3.8+ from https://www.python.org/downloads/
    pause
    exit /b 1
)
echo ✅ Python installed

:: Check Outlook
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ Microsoft Outlook not installed
    echo Please install Microsoft Outlook 2016+ before continuing
    pause
    exit /b 1
)
echo ✅ Microsoft Outlook installed

echo.
echo Step 2: Installing dependencies...
pip install flask flask-cors pywin32 google-generativeai requests >nul 2>&1
echo ✅ Dependencies installed

echo.
echo Step 3: Creating shortcuts...
echo.

:: Create desktop shortcut
echo [InternetShortcut] > "%USERPROFILE%\Desktop\AI Email Manager.url"
echo URL=file:///%CD%/outlook_email_analyzer.py >> "%USERPROFILE%\Desktop\AI Email Manager.url"
echo IconFile=%CD%/outlook_addin/icon-32.ico >> "%USERPROFILE%\Desktop\AI Email Manager.url"
echo IconIndex=0 >> "%USERPROFILE%\Desktop\AI Email Manager.url"

:: Create start menu shortcut
echo [InternetShortcut] > "%APPDATA%\Microsoft\Windows\Start Menu\Programs\AI Email Manager.url"
echo URL=file:///%CD%/outlook_email_analyzer.py >> "%APPDATA%\Microsoft\Windows\Start Menu\Programs\AI Email Manager.url"
echo IconFile=%CD%/outlook_addin/icon-32.ico >> "%APPDATA%\Microsoft\Windows\Start Menu\Programs\AI Email Manager.url"
echo IconIndex=0 >> "%APPDATA%\Microsoft\Windows\Start Menu\Programs\AI Email Manager.url"

echo ✅ Shortcuts created

echo.
echo ========================================
echo           Installation Complete!
echo ========================================
echo.
echo 📋 How to use:
echo.
echo 1. 🔧 Start the main server:
echo    python app_with_ai.py
echo.
echo 2. 📧 Open Outlook and select an email
echo.
echo 3. 🚀 Run the analyzer:
echo    python outlook_email_analyzer.py
echo    OR click "AI Email Manager" shortcut
echo.
echo 4. 🎯 The analyzer will:
echo    - Connect to Outlook
echo    - Get the selected email
echo    - Analyze it with AI
echo    - Show the results
echo.
echo 📞 Support:
echo - Make sure the server is running on localhost:5000
echo - Select an email in Outlook before running the analyzer
echo - Check for any error messages
echo.
echo Press any key to close...
pause >nul




