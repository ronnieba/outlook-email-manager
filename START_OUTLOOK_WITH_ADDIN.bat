@echo off
echo ======================================
echo Starting Outlook with AI Email Manager Add-in
echo ======================================
echo.

REM Kill existing Outlook processes
echo Closing existing Outlook processes...
taskkill /f /im OUTLOOK.EXE 2>nul
timeout /t 2 >nul

echo.
echo Starting Outlook...
start "" "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"

echo.
echo ======================================
echo Outlook started!
echo The Add-in should load automatically.
echo ======================================
echo.
echo Press any key to exit...
pause >nul

