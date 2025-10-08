@echo off
chcp 65001 > nul
setlocal

echo.
echo  ================================================================
echo      בדיקת תוסף COM אולטרה-פשוט
echo  ================================================================
echo.

:: ניקוי קבצי בדיקה ישנים
echo [1] ניקוי קבצי בדיקה ישנים...
del "%TEMP%\ultra_simple_*.txt" >nul 2>&1
echo   [✓] ניקוי הושלם

:: בדיקת Python
echo [2] בדיקת Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] Python לא מותקן
    goto:end
)
echo   [✓] Python מותקן

:: בדיקת pywin32
echo [3] בדיקת pywin32...
python -c "import win32com.client; print('pywin32 OK')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] pywin32 לא מותקן
    goto:end
)
echo   [✓] pywin32 מותקן

:: בדיקת קובץ התוסף
echo [4] בדיקת קובץ התוסף...
if not exist "ultra_simple_addin.py" (
    echo   [❌] קובץ התוסף לא נמצא
    goto:end
)
echo   [✓] קובץ התוסף קיים

:: ביטול רישום קודם
echo [5] ביטול רישום קודם...
python ultra_simple_addin.py --unregister >nul 2>&1
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\UltraSimpleAddin.Addin" /f >nul 2>&1
echo   [✓] ביטול רישום הושלם

:: רישום התוסף
echo [6] רישום התוסף...
python ultra_simple_addin.py --register
if %errorlevel% neq 0 (
    echo   [❌] רישום נכשל
    goto:end
)
echo   [✓] רישום הושלם

:: הוספה ל-Registry
echo [7] הוספה ל-Registry...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\UltraSimpleAddin.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\UltraSimpleAddin.Addin" /v "FriendlyName" /t REG_SZ /d "Ultra Simple Addin" /f >nul
if %errorlevel% neq 0 (
    echo   [❌] הוספה ל-Registry נכשלה
    goto:end
)
echo   [✓] הוספה ל-Registry הושלמה

:: בדיקת יצירת אובייקט
echo [8] בדיקת יצירת אובייקט...
python -c "import win32com.client; obj = win32com.client.Dispatch('UltraSimpleAddin.Addin'); print('Object created')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] לא ניתן ליצור אובייקט
    goto:end
)
echo   [✓] אובייקט נוצר בהצלחה

:: בדיקת קבצי בדיקה
echo [9] בדיקת קבצי בדיקה...
if exist "%TEMP%\ultra_simple_init.txt" (
    echo   [✓] קובץ אתחול נוצר
) else (
    echo   [⚠️] קובץ אתחול לא נוצר
)

:: בדיקת Registry
echo [10] בדיקת Registry...
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\UltraSimpleAddin.Addin" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] Registry לא נמצא
    goto:end
)
echo   [✓] Registry נמצא

echo.
echo  ================================================================
echo                      בדיקה הושלמה! 🎉
echo  ================================================================
echo.
echo  התוסף UltraSimpleAddin מותקן ומוכן!
echo.
echo  מה לעשות עכשיו:
echo  1. פתח את Microsoft Outlook
echo  2. לך ל-File ^> Options ^> Add-ins
echo  3. בדוק ש-"Ultra Simple Addin" מופיע ברשימה
echo  4. ודא שהוא מסומן ב-V (מופעל)
echo.
echo  5. אם התוסף נטען בהצלחה:
echo     - בדוק את הקבצים ב: %TEMP%\ultra_simple_*.txt
echo     - אמור להופיע: ultra_simple_connected.txt
echo     - ואחר כך: ultra_simple_startup.txt
echo.
echo  6. אם עדיין יש שגיאה:
echo     - בדוק את Event Viewer של Windows
echo     - חפש שגיאות ב-Outlook
echo.

:end
echo לחץ על כל מקש לסגירה...
pause > nul
endlocal


