@echo off
echo ========================================
echo התקנת תוסף COM אמיתי עם קובץ EXE
echo ========================================

echo.
echo 1. מבטל רישום תוספים קודמים...
python -c "import win32com.server.register; win32com.server.register.UnregisterServer('{CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC}')" 2>nul

echo.
echo 2. רושם את התוסף החדש...
python real_com_addin.py

echo.
echo 3. מעדכן את רישום Outlook...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\RealCOMAddin.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\RealCOMAddin.Addin" /v "FriendlyName" /t REG_SZ /d "Real COM Addin" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\RealCOMAddin.Addin" /v "Description" /t REG_SZ /d "Real COM Addin for Testing" /f

echo.
echo 4. מוסיף הרשאות אימות...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\RealCOMAddin.Addin" /v "Authentication" /t REG_DWORD /d 1 /f

echo.
echo 5. מוסיף נתיב לקובץ EXE...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\RealCOMAddin.Addin" /v "Path" /t REG_SZ /d "%CD%\real_com_addin.exe" /f

echo.
echo 6. בודק את הרישום...
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\RealCOMAddin.Addin"

echo.
echo ========================================
echo ההתקנה הושלמה!
echo ========================================
echo.
echo עכשיו:
echo 1. סגור את Outlook לחלוטין
echo 2. פתח את Outlook מחדש
echo 3. בדוק אם התוסף נטען
echo 4. בדוק אם נוצרו קבצי טסט בתיקיית TEMP
echo.
pause


