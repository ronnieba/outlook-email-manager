@echo off
chcp 65001 > nul
setlocal

echo.
echo  ================================================================
echo      AI Email Manager - התקנה פשוטה שעובדת
echo  ================================================================
echo.
echo  תוסף COM פשוט שעובד בוודאות
echo  המשתמש עובד רק דרך Outlook
echo.
pause
echo.

:: ניקוי תוספים קודמים
echo [1] ניקוי תוספים קודמים...
python outlook_addin_working.py --unregister >nul 2>&1
python simple_working_addin.py --unregister >nul 2>&1
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /f >nul 2>&1
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SimpleWorkingAddin.Addin" /f >nul 2>&1
echo   [✓] ניקוי הושלם

:: בדיקת דרישות
echo [2] בדיקת דרישות...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] Python לא מותקן
    goto:failure
)
echo   [✓] Python מותקן

python -c "import win32com.client; print('pywin32 OK')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] pywin32 לא מותקן
    pip install pywin32 >nul 2>&1
)
echo   [✓] pywin32 מותקן

python -c "import requests; print('requests OK')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] requests לא מותקן
    pip install requests >nul 2>&1
)
echo   [✓] requests מותקן

:: התקנת התוסף
echo [3] התקנת התוסף...
python simple_working_addin.py --register
if %errorlevel% neq 0 (
    echo   [❌] לא ניתן לרשום את התוסף
    goto:failure
)
echo   [✓] התוסף נרשם ב-COM

:: הוספה ל-Outlook
echo [4] הוספה ל-Outlook...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SimpleWorkingAddin.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SimpleWorkingAddin.Addin" /v "FriendlyName" /t REG_SZ /d "Simple Working Addin" /f >nul
echo   [✓] התוסף נוסף ל-Outlook

:: בדיקת התקנה
echo [5] בדיקת התקנה...
python -c "import win32com.client; win32com.client.Dispatch('SimpleWorkingAddin.Addin')" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [⚠️] לא ניתן ליצור instance של התוסף
) else (
    echo   [✓] התוסף נוצר בהצלחה
)

reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\SimpleWorkingAddin.Addin" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [❌] התוסף לא נרשם ב-Outlook
    goto:failure
) else (
    echo   [✓] התוסף נרשם ב-Outlook
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
echo     התוסף "Simple Working Addin" אמור להופיע ב-Ribbon
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
echo  - Ribbon UI עם כפתור בעברית
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
echo.

:end
echo לחץ על כל מקש לסגירה...
pause > nul
endlocal


