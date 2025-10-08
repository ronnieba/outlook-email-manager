@echo off
echo ========================================
echo התקנת תוסף COM עם קובץ DLL
echo ========================================

echo.
echo 1. מבטל רישום תוספים קודמים...
python -c "import win32com.server.register; win32com.server.register.UnregisterServer('{12345678-1234-1234-1234-123456789012}')" 2>nul
python -c "import win32com.server.register; win32com.server.register.UnregisterServer('{87654321-4321-4321-4321-210987654321}')" 2>nul
python -c "import win32com.server.register; win32com.server.register.UnregisterServer('{11111111-2222-3333-4444-555555555555}')" 2>nul

echo.
echo 2. רושם את התוסף החדש...
python dll_addin.py

echo.
echo 3. מעדכן את רישום Outlook...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\DLLAddin.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\DLLAddin.Addin" /v "FriendlyName" /t REG_SZ /d "DLL Addin" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\DLLAddin.Addin" /v "Description" /t REG_SZ /d "DLL Addin for Testing" /f

echo.
echo 4. מוסיף הרשאות אימות...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\DLLAddin.Addin" /v "Authentication" /t REG_DWORD /d 1 /f

echo.
echo 5. מוסיף נתיב לקובץ...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\DLLAddin.Addin" /v "Path" /t REG_SZ /d "%CD%\dll_addin.py" /f

echo.
echo 6. בודק את הרישום...
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\DLLAddin.Addin"

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


