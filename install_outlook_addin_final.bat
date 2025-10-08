@echo off
chcp 65001 > nul
setlocal

echo.
echo  ================================================================
echo      AI Email Manager - התקנה שעובדת בוודאות
echo  ================================================================
echo.
echo  תוסף COM עם Ribbon UI שעובד ישירות מתוך Outlook
echo  המשתמש עובד רק דרך Outlook - לא צריך מסוף Python
echo.
pause
echo.

:: -------------------------------------------------
:: שלב 1: ניקוי מוחלט
:: -------------------------------------------------
echo [שלב 1/5] ניקוי מוחלט...

echo   - מבטל רישום תוספים קודמים...
python outlook_addin_working.py --unregister >nul 2>&1
python outlook_com_addin_final.py --unregister >nul 2>&1
python working_outlook_addin.py --unregister >nul 2>&1
python ultra_simple_addin.py --unregister >nul 2>&1

echo   - מוחק רישומים ישנים ב-Registry...
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /f >nul 2>&1
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\UltraSimpleAddin.Addin" /f >nul 2>&1
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\WorkingAIEmailManager.Addin" /f >nul 2>&1

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

:: בדיקת requests
python -c "import requests; print('requests OK')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] requests לא מותקן
    echo   מתקין requests...
    pip install requests >nul 2>&1
    if %errorlevel% neq 0 (
        echo   [❌] לא ניתן להתקין requests
        goto:failure
    )
)
echo   [✓] requests מותקן

:: בדיקת Outlook
python -c "import win32com.client; win32com.client.Dispatch('Outlook.Application')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [⚠️] Outlook לא פתוח או לא מותקן
    echo   ודא ש-Outlook פתוח לפני השימוש
) else (
    echo   [✓] Outlook זמין
)
echo.

:: -------------------------------------------------
:: שלב 3: התקנת התוסף
:: -------------------------------------------------
echo [שלב 3/5] התקנת התוסף...

:: בדיקה שהקובץ קיים
if not exist "outlook_addin_working.py" (
    echo   [❌] קובץ התוסף לא נמצא: outlook_addin_working.py
    goto:failure
)
echo   [✓] קובץ התוסף קיים

:: רישום התוסף
echo   - רושם את התוסף ב-COM...
python outlook_addin_working.py --register
if %errorlevel% neq 0 (
    echo   [❌] לא ניתן לרשום את התוסף
    echo   נסה להפעיל את הסקריפט כמנהל
    goto:failure
)
echo   [✓] התוסף נרשם ב-COM

:: הוספה ל-Outlook
echo   - מוסיף את התוסף ל-Outlook...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "FriendlyName" /t REG_SZ /d "AI Email Manager" /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "Description" /t REG_SZ /d "AI-powered email analysis for Outlook" /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "CommandLineSafe" /t REG_DWORD /d 0 /f >nul

echo   [✓] התוסף נוסף ל-Outlook
echo.

:: -------------------------------------------------
:: שלב 4: בדיקת ההתקנה
:: -------------------------------------------------
echo [שלב 4/5] בדיקת ההתקנה...

:: בדיקת רישום COM
python -c "import win32com.client; win32com.client.Dispatch('AIEmailManager.Addin')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [⚠️] לא ניתן ליצור instance של התוסף
) else (
    echo   [✓] התוסף נוצר בהצלחה
)

:: בדיקת רישום Outlook
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] התוסף לא נרשם ב-Outlook
    goto:failure
) else (
    echo   [✓] התוסף נרשם ב-Outlook
)

:: בדיקת השרת
echo [שלב 5/5] בדיקת השרת...
python -c "import requests; requests.get('http://localhost:5000/api/status', timeout=2)" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [⚠️] השרת לא פועל
    echo   הפעל: python app_with_ai.py
) else (
    echo   [✓] השרת פועל
)

echo.
echo  ================================================================
echo                      התקנה הושלמה בהצלחה! 🎉
echo  ================================================================
echo.
echo  מה לעשות עכשיו:
echo.
echo  1. הפעל את השרת הראשי:
echo     ^> python app_with_ai.py
echo.
echo  2. פתח את Microsoft Outlook
echo     התוסף "AI Email Manager" אמור להופיע ב-Ribbon
echo.
echo  3. לבדיקה:
echo     - בחר מייל ב-Outlook
echo     - לחץ על Tab "AI Email Manager"
echo     - לחץ על "נתח מייל נוכחי"
echo.
echo  4. אם התוסף לא מופיע:
echo     - סגור את Outlook לחלוטין
echo     - הפעל מחדש את הסקריפט
echo     - פתח את Outlook שוב
echo.
echo  יתרונות:
echo  - המשתמש עובד רק דרך Outlook
echo  - לא צריך מסוף Python
echo  - Ribbon UI עם כפתורים בעברית
echo  - מוסיף Custom Properties למיילים
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
echo  3. בדוק את הלוגים ב: %TEMP%\outlook_addin_working.log
echo.

:end
echo לחץ על כל מקש לסגירה...
pause > nul
endlocal