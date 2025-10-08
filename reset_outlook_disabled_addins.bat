@echo off
chcp 65001 > nul
echo ================================
echo איפוס רשימת תוספים מושבתים
echo ================================
echo.

echo שלב 1: סגירת Outlook...
taskkill /F /IM OUTLOOK.EXE 2>nul
timeout /t 2 /nobreak >nul

echo.
echo שלב 2: מחיקת רשימת תוספים מושבתים...

REM מחיקת Resiliency Keys
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems" /f 2>nul
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\NotificationResilientComponents" /f 2>nul
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList" /f 2>nul
reg delete "HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList" /f 2>nul

REM וגם עבור Office 365 (גרסה 16.0)
reg delete "HKCU\Software\Microsoft\Office\Outlook\Addins\AIEmailManagerAddin\DisabledItems" /f 2>nul

echo ✓ רשימת תוספים מושבתים נמחקה!

echo.
echo שלב 3: ודא שהתוסף מופעל ב-Registry...
reg add "HKCU\Software\Microsoft\Office\Outlook\Addins\AIEmailManagerAddin" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul

echo ✓ LoadBehavior הוגדר ל-3 (טעון תמיד)

echo.
echo ================================
echo ✓ הסתיים!
echo ================================
echo.
echo עכשיו פתח את Outlook ובדוק שוב
echo.
pause
