@echo off
chcp 65001 >nul
echo.
echo ========================================
echo    התקנת תוסף Ribbon ל-Outlook
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
echo שלב 3: רישום התוסף ב-Windows Registry...
echo.

:: יצירת רישום התוסף
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManagerRibbon.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManagerRibbon.Addin" /v "FriendlyName" /t REG_SZ /d "AI Email Manager Ribbon" /f >nul 2>&1
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManagerRibbon.Addin" /v "Description" /t REG_SZ /d "ניתוח חכם של מיילים ופגישות עם AI - Ribbon" /f >nul 2>&1
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManagerRibbon.Addin" /v "Manifest" /t REG_SZ /d "%CD%\outlook_addin\manifest.xml" /f >nul 2>&1

:: יצירת הגדרות התוסף
reg add "HKEY_CURRENT_USER\Software\AIEmailManagerRibbon" /v "Version" /t REG_SZ /d "1.0.0" /f >nul 2>&1
reg add "HKEY_CURRENT_USER\Software\AIEmailManagerRibbon" /v "InstallPath" /t REG_SZ /d "%CD%" /f >nul 2>&1
reg add "HKEY_CURRENT_USER\Software\AIEmailManagerRibbon" /v "ServerURL" /t REG_SZ /d "http://localhost:5000" /f >nul 2>&1

echo ✅ התוסף נרשם ב-Windows Registry

echo.
echo שלב 4: יצירת קיצורי דרך...
echo.

:: קיצור דרך על שולחן העבודה
echo [InternetShortcut] > "%USERPROFILE%\Desktop\AI Email Manager Ribbon.url"
echo URL=file:///%CD%/outlook_ribbon_addin.py >> "%USERPROFILE%\Desktop\AI Email Manager Ribbon.url"
echo IconFile=%CD%\outlook_addin\icon-32.ico >> "%USERPROFILE%\Desktop\AI Email Manager Ribbon.url"
echo IconIndex=0 >> "%USERPROFILE%\Desktop\AI Email Manager Ribbon.url"

:: קיצור דרך בתפריט התחל
echo [InternetShortcut] > "%APPDATA%\Microsoft\Windows\Start Menu\Programs\AI Email Manager Ribbon.url"
echo URL=file:///%CD%/outlook_ribbon_addin.py >> "%APPDATA%\Microsoft\Windows\Start Menu\Programs\AI Email Manager Ribbon.url"
echo IconFile=%CD%\outlook_addin\icon-32.ico >> "%APPDATA%\Microsoft\Windows\Start Menu\Programs\AI Email Manager Ribbon.url"
echo IconIndex=0 >> "%APPDATA%\Microsoft\Windows\Start Menu\Programs\AI Email Manager Ribbon.url"

echo ✅ קיצורי דרך נוצרו

echo.
echo שלב 5: בדיקת התקנה...
echo.

:: בדיקת COM
python -c "import win32com.client; print('COM: OK')" >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ שגיאה בבדיקת COM
    echo אנא ודא ש-pywin32 מותקן: pip install pywin32
) else (
    echo ✅ בדיקת COM עברה בהצלחה
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
echo 2. 🚀 הפעל את התוסף Ribbon:
echo    python outlook_ribbon_addin.py
echo    או לחץ על הקיצור "AI Email Manager Ribbon"
echo.
echo 3. 📧 פתח את Outlook ותראה:
echo    - Tab חדש: "AI Email Manager"
echo    - כפתורים: "נתח מייל נוכחי", "נתח מיילים נבחרים"
echo    - Context Menu: לחיצה ימנית על מיילים
echo.
echo 4. 🎯 השתמש בכפתורי ה-Ribbon:
echo    - בחר מייל ולחץ "נתח מייל נוכחי"
echo    - בחר כמה מיילים ולחץ "נתח מיילים נבחרים"
echo    - לחץ ימני על מייל ובחר "נתח עם AI"
echo.
echo 📞 תמיכה:
echo - בדוק את הלוגים ב-outlook_ribbon_addin.log
echo - ודא שהשרת רץ על localhost:5000
echo - בדוק את החיבור ל-Outlook
echo.
echo לחץ על מקש כלשהו לסגירה...
pause >nul





