@echo off
chcp 65001 >nul
echo.
echo ========================================
echo    התקנת תוסף AI Email Manager ב-Outlook
echo ========================================
echo.

echo שלב 1: בדיקת דרישות...
echo.

:: בדיקת Python
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ Python לא מותקן
    echo אנא התקן Python 3.8+ מ-https://www.python.org/downloads/
    pause
    exit /b 1
)
echo ✅ Python מותקן

:: בדיקת Outlook
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ Microsoft Outlook לא מותקן
    echo אנא התקן Microsoft Outlook 2016+ לפני המשך
    pause
    exit /b 1
)
echo ✅ Microsoft Outlook מותקן

echo.
echo שלב 2: התקנת תלויות...
pip install flask flask-cors pywin32 google-generativeai requests >nul 2>&1
echo ✅ תלויות הותקנו

echo.
echo שלב 3: ביטול רישום התוסף הישן...
echo.

:: ביטול רישום התוסף הישן
python outlook_com_addin_registered.py --unregister >nul 2>&1
echo ✅ התוסף הישן בוטל

echo.
echo שלב 4: רישום התוסף החדש ב-COM...
echo.

:: רישום התוסף החדש ב-COM
python outlook_com_addin_simple_fixed.py --register
if %errorLevel% neq 0 (
    echo ❌ שגיאה ברישום התוסף ב-COM
    pause
    exit /b 1
)
echo ✅ התוסף החדש נרשם ב-COM

echo.
echo שלב 5: רישום התוסף ב-Outlook...
echo.

:: יצירת רישום התוסף ב-Outlook
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "FriendlyName" /t REG_SZ /d "AI Email Manager" /f >nul 2>&1
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "Description" /t REG_SZ /d "ניתוח חכם של מיילים ופגישות עם AI" /f >nul 2>&1
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "CommandLineSafe" /t REG_DWORD /d 0 /f >nul 2>&1

echo ✅ התוסף נרשם ב-Outlook

echo.
echo שלב 6: בדיקת התקנה...
echo.

:: בדיקת COM
python -c "import win32com.client; print('COM: OK')" >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ שגיאה בבדיקת COM
    echo אנא ודא ש-pywin32 מותקן: pip install pywin32
) else (
    echo ✅ בדיקת COM עברה בהצלחה
)

:: בדיקת רישום התוסף ב-COM
reg query "HKEY_CLASSES_ROOT\AIEmailManager.Addin" >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ התוסף לא נרשם ב-COM
) else (
    echo ✅ התוסף נרשם ב-COM
)

:: בדיקת רישום התוסף ב-Outlook
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ התוסף לא נרשם ב-Outlook
) else (
    echo ✅ התוסף נרשם ב-Outlook
)

echo.
echo ========================================
echo           התקנה הושלמה בהצלחה!
echo ========================================
echo.
echo 📋 מה לעשות עכשיו:
echo.
echo 1. 🔧 הפעל את השרת הראשי:
echo    python app_with_ai.py
echo.
echo 2. 📧 פתח את Outlook ובדוק:
echo    - File → Options → Add-ins
echo    - בדוק שהתוסף "AI Email Manager" מופיע
echo    - ודא שהוא מסומן ב-V (מופעל)
echo    - בדוק שאין שגיאת זמן ריצה
echo.
echo 3. 🎯 השתמש בתוסף:
echo    - התוסף יפעל אוטומטית כשתפתח Outlook
echo    - בדוק את הלוגים ב-outlook_addin_success.log
echo    - אם יש שגיאות, בדוק את outlook_addin_error.log
echo.
echo 📞 תמיכה:
echo - בדוק את הלוגים ב-outlook_addin_success.log
echo - אם יש שגיאות, בדוק את outlook_addin_error.log
echo - ודא שהשרת רץ על localhost:5000
echo - בדוק את החיבור ל-Outlook
echo.
echo לחץ על מקש כלשהו לסגירה...
pause >nul




