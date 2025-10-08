@echo off
echo ========================================
echo התקנת תוסף COM פשוט
echo ========================================

echo.
echo 1. מבטל רישום תוספים קודמים...
reg delete "HKEY_CLASSES_ROOT\CLSID\{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}" /f 2>nul
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SimpleCOMAddin.Addin" /f 2>nul

echo.
echo 2. רושם את התוסף החדש...
python simple_com_addin.py

echo.
echo 3. מעדכן את רישום Outlook...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SimpleCOMAddin.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SimpleCOMAddin.Addin" /v "FriendlyName" /t REG_SZ /d "Simple COM Addin" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SimpleCOMAddin.Addin" /v "Description" /t REG_SZ /d "Simple COM Addin for Testing" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SimpleCOMAddin.Addin" /v "Authentication" /t REG_DWORD /d 1 /f

echo.
echo 4. בודק את הרישום...
reg query "HKEY_CLASSES_ROOT\CLSID\{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}"
echo.
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SimpleCOMAddin.Addin"

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