@echo off
chcp 65001 > nul
setlocal

echo.
echo  ================================================================
echo      AI Email Manager - פתרון שעובד בוודאות
echo  ================================================================
echo.
echo  במקום תוסף COM מורכב, פתרון פשוט שעובד!
echo.
echo  מה זה עושה:
echo  - מתחבר ל-Outlook ישירות
echo  - מנתח מיילים עם AI
echo  - מוסיף Custom Properties למיילים
echo  - עובד בוודאות ללא בעיות COM
echo.
pause
echo.

:: בדיקת Python
echo [1] בדיקת Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] Python לא מותקן
    goto:end
)
echo   [✓] Python מותקן

:: בדיקת pywin32
echo [2] בדיקת pywin32...
python -c "import win32com.client; print('pywin32 OK')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] pywin32 לא מותקן
    echo   מתקין pywin32...
    pip install pywin32 >nul 2>&1
    if %errorlevel% neq 0 (
        echo   [❌] לא ניתן להתקין pywin32
        goto:end
    )
)
echo   [✓] pywin32 מותקן

:: בדיקת requests
echo [3] בדיקת requests...
python -c "import requests; print('requests OK')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] requests לא מותקן
    echo   מתקין requests...
    pip install requests >nul 2>&1
    if %errorlevel% neq 0 (
        echo   [❌] לא ניתן להתקין requests
        goto:end
    )
)
echo   [✓] requests מותקן

:: בדיקת קובץ המנתח
echo [4] בדיקת קובץ המנתח...
if not exist "simple_email_analyzer.py" (
    echo   [❌] קובץ המנתח לא נמצא
    goto:end
)
echo   [✓] קובץ המנתח קיים

:: בדיקת Outlook
echo [5] בדיקת Outlook...
python -c "import win32com.client; win32com.client.Dispatch('Outlook.Application')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [⚠️] Outlook לא פתוח או לא מותקן
    echo   ודא ש-Outlook פתוח לפני השימוש
) else (
    echo   [✓] Outlook זמין
)

:: בדיקת השרת
echo [6] בדיקת השרת...
python -c "import requests; requests.get('http://localhost:5000/api/status', timeout=2)" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [⚠️] השרת לא פועל
    echo   הפעל: python app_with_ai.py
) else (
    echo   [✓] השרת פועל
)

echo.
echo  ================================================================
echo                      הכל מוכן! 🎉
echo  ================================================================
echo.
echo  איך להשתמש:
echo.
echo  1. ודא ש-Outlook פתוח
echo  2. ודא שהשרת פועל: python app_with_ai.py
echo  3. הפעל את המנתח: python simple_email_analyzer.py
echo.
echo  4. במנתח:
echo     - בחר מייל ב-Outlook
echo     - לחץ 1 לניתוח המייל הנוכחי
echo     - לחץ 2 לניתוח כל המיילים הנבחרים
echo.
echo  5. הניתוח יופיע במסוף ויוסף Custom Properties למייל
echo.
echo  יתרונות:
echo  - עובד בוודאות ללא בעיות COM
echo  - פשוט לשימוש
echo  - לא דורש התקנה מורכבת
echo  - מוסיף Custom Properties למיילים
echo.

:end
echo לחץ על כל מקש לסגירה...
pause > nul
endlocal


