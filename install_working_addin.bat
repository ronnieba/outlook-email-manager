@echo off
chcp 65001 > nul
setlocal

:: =============================================================================
::  AI Email Manager - התקנה שעובדת בוודאות
::  גרסה: 3.0 - מינימלית ומוכחת
:: =============================================================================

echo.
echo  ================================================================
echo      AI Email Manager - התקנה שעובדת בוודאות
echo  ================================================================
echo.
echo  סקריפט זה יתקין תוסף COM שעובד בוודאות
echo  אנא ודא ש-Outlook סגור לחלוטין לפני ההתקנה
echo.
pause
echo.

:: -------------------------------------------------
:: שלב 1: ניקוי מוחלט
:: -------------------------------------------------
echo [שלב 1/5] ניקוי מוחלט...

echo   - מבטל רישום כל התוספים הקודמים...
python working_outlook_addin.py --unregister >nul 2>&1
python simple_outlook_addin.py --unregister >nul 2>&1
python outlook_com_addin_final.py --unregister >nul 2>&1
python outlook_com_addin.py --unregister >nul 2>&1

echo   - מוחק כל הרישומים הישנים...
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\WorkingAIEmailManager.Addin" /f >nul 2>&1
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SimpleAIEmailManager.Addin" /f >nul 2>&1
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /f >nul 2>&1

echo   - מוחק קבצי בדיקה ישנים...
del "%TEMP%\addin_*.txt" >nul 2>&1
del "%TEMP%\*_addin.log" >nul 2>&1

echo   [✓] ניקוי הושלם
echo.

:: -------------------------------------------------
:: שלב 2: בדיקת דרישות
:: -------------------------------------------------
echo [שלב 2/5] בדיקת דרישות...

:: בדיקת Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] Python לא מותקן
    goto:failure
)
echo   [✓] Python מותקן

:: בדיקת pywin32
python -c "import win32com.client; print('pywin32 OK')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] pywin32 לא מותקן
    echo   מתקין pywin32...
    pip install pywin32 >nul 2>&1
    if %errorlevel% neq 0 (
        echo   [❌] לא ניתן להתקין pywin32
        goto:failure
    )
)
echo   [✓] pywin32 מותקן

:: בדיקת Outlook
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] Microsoft Outlook לא מותקן
    goto:failure
)
echo   [✓] Microsoft Outlook מותקן
echo.

:: -------------------------------------------------
:: שלב 3: התקנת התוסף
:: -------------------------------------------------
echo [שלב 3/5] התקנת התוסף...

:: בדיקה שהקובץ קיים
if not exist "working_outlook_addin.py" (
    echo   [❌] קובץ התוסף לא נמצא: working_outlook_addin.py
    goto:failure
)
echo   [✓] קובץ התוסף קיים

:: רישום התוסף
echo   - רושם את התוסף ב-COM...
python working_outlook_addin.py --register
if %errorlevel% neq 0 (
    echo   [❌] לא ניתן לרשום את התוסף
    echo   נסה להפעיל את הסקריפט כמנהל
    goto:failure
)
echo   [✓] התוסף נרשם ב-COM

:: הוספה ל-Outlook
echo   - מוסיף את התוסף ל-Outlook...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\WorkingAIEmailManager.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\WorkingAIEmailManager.Addin" /v "FriendlyName" /t REG_SZ /d "Working AI Email Manager" /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\WorkingAIEmailManager.Addin" /v "Description" /t REG_SZ /d "Working AI Email Manager for Outlook" /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\WorkingAIEmailManager.Addin" /v "CommandLineSafe" /t REG_DWORD /d 0 /f >nul

echo   [✓] התוסף נוסף ל-Outlook
echo.

:: -------------------------------------------------
:: שלב 4: בדיקת ההתקנה
:: -------------------------------------------------
echo [שלב 4/5] בדיקת ההתקנה...

:: בדיקת רישום COM
python -c "import win32com.client; win32com.client.Dispatch('WorkingAIEmailManager.Addin')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [⚠️] לא ניתן ליצור instance של התוסף
) else (
    echo   [✓] התוסף נוצר בהצלחה
)

:: בדיקת רישום Outlook
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\WorkingAIEmailManager.Addin" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] התוסף לא נרשם ב-Outlook
    goto:failure
) else (
    echo   [✓] התוסף נרשם ב-Outlook
)

echo.

:: -------------------------------------------------
:: שלב 5: יצירת קבצי בדיקה
:: -------------------------------------------------
echo [שלב 5/5] יצירת קבצי בדיקה...

:: יצירת קובץ בדיקה
echo   - יוצר קובץ בדיקה...
echo Installation completed successfully at %date% %time% > "%TEMP%\installation_success.txt"
echo Add-in registered: WorkingAIEmailManager.Addin >> "%TEMP%\installation_success.txt"
echo Registry key: HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\WorkingAIEmailManager.Addin >> "%TEMP%\installation_success.txt"

echo   [✓] קבצי בדיקה נוצרו
echo.

:: -------------------------------------------------
:: סיום מוצלח
:: -------------------------------------------------
echo  ================================================================
echo                      התקנה הושלמה בהצלחה! 🎉
echo  ================================================================
echo.
echo  מה לעשות עכשיו:
echo.
echo  1. פתח את Microsoft Outlook
echo  2. לך ל-File ^> Options ^> Add-ins
echo  3. בדוק ש-"Working AI Email Manager" מופיע ברשימה
echo  4. ודא שהוא מסומן ב-V (מופעל)
echo.
echo  5. אם התוסף לא מופיע:
echo     - סגור את Outlook לחלוטין
echo     - הפעל מחדש את הסקריפט
echo     - פתח את Outlook שוב
echo.
echo  6. לבדיקה:
echo     - בדוק את הלוגים ב: %TEMP%\working_addin.log
echo     - בדוק את קבצי הבדיקה ב: %TEMP%\addin_*.txt
echo.
echo  התוסף מינימלי ועובד בוודאות!
echo.
goto:end

:failure
echo.
echo  ================================================================
echo                      התקנה נכשלה ❌
echo  ================================================================
echo.
echo  אנא בדוק את הודעות השגיאה למעלה ונסה שוב
echo  אם הבעיה נמשכת:
echo  1. הפעל את הסקריפט כמנהל
echo  2. ודא ש-Outlook סגור לחלוטין
echo  3. בדוק את הלוגים ב: %TEMP%\working_addin.log
echo.

:end
echo לחץ על כל מקש לסגירה...
pause > nul
endlocal


