@echo off
echo ========================================
echo התקנת תוסף COM עם קובץ EXE אמיתי - גרסה 2
echo ========================================

echo.
echo 1. מבטל רישום תוספים קודמים...
reg delete "HKEY_CLASSES_ROOT\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}" /f 2>nul
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EXEAddin.Addin" /f 2>nul

echo.
echo 2. יוצר רישום COM ידני לקובץ EXE...
reg add "HKEY_CLASSES_ROOT\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}" /v "" /t REG_SZ /d "EXE Addin" /f
reg add "HKEY_CLASSES_ROOT\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}\ProgID" /v "" /t REG_SZ /d "EXEAddin.Addin" /f
reg add "HKEY_CLASSES_ROOT\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}\LocalServer32" /v "" /t REG_SZ /d "%CD%\dist\exe_addin.exe" /f
reg add "HKEY_CLASSES_ROOT\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}\Implemented Categories\{B3EF80D0-68E2-11D0-A689-00C04FD658FF}" /f

echo.
echo 3. יוצר רישום ProgID...
reg add "HKEY_CLASSES_ROOT\EXEAddin.Addin" /v "" /t REG_SZ /d "EXE Addin" /f
reg add "HKEY_CLASSES_ROOT\EXEAddin.Addin\CLSID" /v "" /t REG_SZ /d "{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}" /f

echo.
echo 4. מעדכן את רישום Outlook...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EXEAddin.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EXEAddin.Addin" /v "FriendlyName" /t REG_SZ /d "EXE Addin" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EXEAddin.Addin" /v "Description" /t REG_SZ /d "EXE Addin for Testing" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EXEAddin.Addin" /v "Authentication" /t REG_DWORD /d 1 /f

echo.
echo 5. בודק את הרישום...
reg query "HKEY_CLASSES_ROOT\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}"
echo.
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\EXEAddin.Addin"

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


