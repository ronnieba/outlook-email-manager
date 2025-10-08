@echo off
chcp 65001 > nul
setlocal

echo.
echo  ================================================================
echo      AI Email Manager - בדיקת התקנה
echo  ================================================================
echo.

:: בדיקת Python
echo [1] בדיקת Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   ❌ Python לא מותקן
    goto:end
) else (
    echo   ✅ Python מותקן
)

:: בדיקת תלויות
echo [2] בדיקת תלויות...
python -c "import win32com.client; print('pywin32: OK')" 2>nul
if %errorlevel% neq 0 (
    echo   ❌ pywin32 לא מותקן
    goto:end
) else (
    echo   ✅ pywin32 מותקן
)

python -c "import requests; print('requests: OK')" 2>nul
if %errorlevel% neq 0 (
    echo   ❌ requests לא מותקן
    goto:end
) else (
    echo   ✅ requests מותקן
)

:: בדיקת קובץ התוסף
echo [3] בדיקת קובץ התוסף...
if not exist "outlook_com_addin_final.py" (
    echo   ❌ קובץ התוסף לא נמצא
    goto:end
) else (
    echo   ✅ קובץ התוסף קיים
)

:: בדיקת רישום COM
echo [4] בדיקת רישום COM...
python outlook_com_addin_final.py --unregister >nul 2>&1
python outlook_com_addin_final.py --register >nul 2>&1
if %errorlevel% neq 0 (
    echo   ❌ רישום COM נכשל
    goto:end
) else (
    echo   ✅ רישום COM הצליח
)

:: בדיקת רישום Outlook
echo [5] בדיקת רישום Outlook...
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" >nul 2>&1
if %errorlevel% neq 0 (
    echo   ❌ רישום Outlook נכשל
    goto:end
) else (
    echo   ✅ רישום Outlook הצליח
)

:: בדיקת השרת
echo [6] בדיקת השרת...
curl -s http://localhost:5000/api/status >nul 2>&1
if %errorlevel% neq 0 (
    echo   ⚠️  השרת לא פועל (זה בסדר אם לא הפעלת אותו)
) else (
    echo   ✅ השרת פועל
)

echo.
echo  ================================================================
echo                      בדיקה הושלמה! 🎉
echo  ================================================================
echo.
echo  התוסף מותקן ומוכן לשימוש!
echo.
echo  מה לעשות עכשיו:
echo  1. הפעל את השרת: python app_with_ai.py
echo  2. פתח את Outlook
echo  3. חפש Tab "AI Email Manager" ב-Ribbon
echo  4. בחר מייל ולחץ "נתח מייל נוכחי"
echo.

:end
echo לחץ על כל מקש לסגירה...
pause > nul
endlocal


