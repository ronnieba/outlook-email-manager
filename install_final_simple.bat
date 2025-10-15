@echo off
chcp 65001 > nul
setlocal

:: =============================================================================
::  AI Email Manager - התקנה פשוטה ועובדת של תוסף COM
::  גרסה: 2.0 - פשוטה ומתקדמת
:: =============================================================================

echo.
echo  ================================================================
echo      AI Email Manager - התקנה פשוטה של תוסף Outlook
echo  ================================================================
echo.
echo  סקריפט זה יתקין את תוסף AI Email Manager ב-Outlook
echo  אנא ודא ש-Outlook סגור לפני ההתקנה
echo.
pause
echo.

:: -------------------------------------------------
:: שלב 1: בדיקת דרישות מערכת
:: -------------------------------------------------
echo [שלב 1/4] בדיקת דרישות מערכת...

:: בדיקת Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   [שגיאה] Python לא מותקן או לא ב-PATH
    echo   אנא התקן Python 3.8+ מ-https://www.python.org/downloads/
    goto:failure
)
echo   [✓] Python מותקן

:: בדיקת Outlook
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [שגיאה] Microsoft Outlook לא מותקן
    echo   אנא התקן Microsoft Outlook 2016 או חדש יותר
    goto:failure
)
echo   [✓] Microsoft Outlook מותקן
echo.

:: -------------------------------------------------
:: שלב 2: ניקוי גרסאות קודמות
:: -------------------------------------------------
echo [שלב 2/4] ניקוי גרסאות קודמות...

echo   - מבטל רישום תוספים קודמים...
python outlook_com_addin.py --unregister >nul 2>&1
python outlook_com_addin_final.py --unregister >nul 2>&1

echo   - מוחק רישומים ישנים ב-Registry...
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /f >nul 2>&1

echo   [✓] ניקוי הושלם
echo.

:: -------------------------------------------------
:: שלב 3: התקנת תלויות
:: -------------------------------------------------
echo [שלב 3/4] התקנת תלויות Python...

echo   - מתקין pywin32...
pip install --upgrade pywin32 >nul 2>&1
if %errorlevel% neq 0 (
    echo   [שגיאה] לא ניתן להתקין pywin32
    echo   נסה להתקין ידנית: pip install pywin32
    goto:failure
)

echo   - מתקין requests...
pip install --upgrade requests >nul 2>&1
if %errorlevel% neq 0 (
    echo   [שגיאה] לא ניתן להתקין requests
    echo   נסה להתקין ידנית: pip install requests
    goto:failure
)

echo   [✓] תלויות הותקנו בהצלחה
echo.

:: -------------------------------------------------
:: שלב 4: התקנת התוסף
:: -------------------------------------------------
echo [שלב 4/4] התקנת התוסף...

:: בדיקה שהקובץ קיים
if not exist "outlook_com_addin_final.py" (
    echo   [שגיאה] קובץ התוסף לא נמצא: outlook_com_addin_final.py
    goto:failure
)

echo   - רושם את התוסף ב-COM...
python outlook_com_addin_final.py --register
if %errorlevel% neq 0 (
    echo   [שגיאה] לא ניתן לרשום את התוסף
    echo   נסה להפעיל את הסקריפט כמנהל
    goto:failure
)

echo   - מוסיף את התוסף ל-Outlook...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "FriendlyName" /t REG_SZ /d "AI Email Manager" /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "Description" /t REG_SZ /d "AI-powered email analysis for Outlook" /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "CommandLineSafe" /t REG_DWORD /d 0 /f >nul

echo   [✓] התוסף הותקן בהצלחה!
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
echo     - סגור את Outlook
echo     - הפעל מחדש את הסקריפט
echo     - פתח את Outlook שוב
echo.
echo  לוגים נשמרים ב: %TEMP%\ai_email_manager.log
echo.
goto:end

:failure
echo.
echo  ================================================================
echo                      התקנה נכשלה ❌
echo  ================================================================
echo.
echo  אנא בדוק את הודעות השגיאה למעלה ונסה שוב
echo  אם הבעיה נמשכת, נסה להפעיל את הסקריפט כמנהל
echo.

:end
echo לחץ על כל מקש לסגירה...
pause > nul
endlocal








