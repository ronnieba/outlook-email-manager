@echo off
echo ========================================
echo התקנת תוסף COM מינימלי לבדיקה
echo ========================================

echo.
echo 1. מבטל רישום תוספים קודמים...
python -c "import win32com.server.register; win32com.server.register.UnregisterServer('{12345678-1234-1234-1234-123456789012}')" 2>nul
python -c "import win32com.server.register; win32com.server.register.UnregisterServer('{87654321-4321-4321-4321-210987654321}')" 2>nul

echo.
echo 2. רושם את התוסף החדש...
python minimal_test_addin.py

echo.
echo 3. מעדכן את רישום Outlook...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\MinimalTestAddin.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\MinimalTestAddin.Addin" /v "FriendlyName" /t REG_SZ /d "Minimal Test Addin" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\MinimalTestAddin.Addin" /v "Description" /t REG_SZ /d "Minimal Test Addin for Testing" /f

echo.
echo 4. בודק את הרישום...
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\MinimalTestAddin.Addin"

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


