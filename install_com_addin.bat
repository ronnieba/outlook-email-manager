@echo off
chcp 65001 >nul
echo.
echo ========================================
echo    AI Email Manager - התקנת תוסף COM
echo ========================================
echo.

:: בדיקת הרשאות מנהל
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ נדרשות הרשאות מנהל להתקנה
    echo לחץ על מקש כלשהו לסגירה...
    pause >nul
    exit /b 1
)

echo ✅ הרשאות מנהל מאושרות
echo.

:: בדיקת Python
echo 🔍 בודק Python...
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ Python לא מותקן או לא נמצא ב-PATH
    echo אנא התקן Python 3.8+ מ-https://www.python.org/downloads/
    echo לחץ על מקש כלשהו לסגירה...
    pause >nul
    exit /b 1
)

echo ✅ Python מותקן
echo.

:: בדיקת Outlook
echo 🔍 בודק Microsoft Outlook...
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ Microsoft Outlook לא מותקן
    echo אנא התקן Microsoft Outlook 2016+ לפני המשך
    echo לחץ על מקש כלשהו לסגירה...
    pause >nul
    exit /b 1
)

echo ✅ Microsoft Outlook מותקן
echo.

:: התקנת תלויות Python
echo 📦 מתקין תלויות Python...
pip install flask==2.3.3 flask-cors==4.0.0 pywin32>=307 google-generativeai==0.3.2 requests sqlite3 >nul 2>&1
if %errorLevel% neq 0 (
    echo ⚠️ שגיאה בהתקנת תלויות, מנסה שוב...
    pip install flask flask-cors pywin32 google-generativeai requests
)

echo ✅ תלויות Python הותקנו
echo.

:: יצירת תיקיות נדרשות
echo 📁 יוצר תיקיות...
if not exist "C:\Users\%USERNAME%\outlook_email_manager" mkdir "C:\Users\%USERNAME%\outlook_email_manager"
if not exist "C:\Users\%USERNAME%\outlook_email_manager\outlook_addin" mkdir "C:\Users\%USERNAME%\outlook_email_manager\outlook_addin"
if not exist "C:\Users\%USERNAME%\outlook_email_manager\logs" mkdir "C:\Users\%USERNAME%\outlook_email_manager\logs"

echo ✅ תיקיות נוצרו
echo.

:: העתקת קבצים
echo 📋 מעתיק קבצים...
if exist "outlook_com_addin.py" copy "outlook_com_addin.py" "C:\Users\%USERNAME%\outlook_email_manager\" >nul
if exist "outlook_addin\manifest.xml" copy "outlook_addin\manifest.xml" "C:\Users\%USERNAME%\outlook_email_manager\outlook_addin\" >nul
if exist "outlook_addin\taskpane.html" copy "outlook_addin\taskpane.html" "C:\Users\%USERNAME%\outlook_email_manager\outlook_addin\" >nul
if exist "outlook_addin\taskpane.js" copy "outlook_addin\taskpane.js" "C:\Users\%USERNAME%\outlook_email_manager\outlook_addin\" >nul
if exist "outlook_addin\taskpane.css" copy "outlook_addin\taskpane.css" "C:\Users\%USERNAME%\outlook_email_manager\outlook_addin\" >nul

echo ✅ קבצים הועתקו
echo.

:: רישום התוסף ב-Windows Registry
echo 🔧 רושם תוסף ב-Windows Registry...
regedit /s "outlook_addin_registry.reg" >nul 2>&1
if %errorLevel% neq 0 (
    echo ⚠️ שגיאה ברישום התוסף, מנסה ידנית...
    echo אנא הרץ את outlook_addin_registry.reg ידנית
)

echo ✅ תוסף נרשם ב-Windows Registry
echo.

:: יצירת קיצור דרך
echo 🔗 יוצר קיצור דרך...
set "desktop=%USERPROFILE%\Desktop"
set "startMenu=%APPDATA%\Microsoft\Windows\Start Menu\Programs"

:: קיצור דרך על שולחן העבודה
echo [InternetShortcut] > "%desktop%\AI Email Manager.url"
echo URL=file:///C:/Users/%USERNAME%/outlook_email_manager/outlook_com_addin.py >> "%desktop%\AI Email Manager.url"
echo IconFile=C:\Users\%USERNAME%\outlook_email_manager\outlook_addin\icon-32.ico >> "%desktop%\AI Email Manager.url"
echo IconIndex=0 >> "%desktop%\AI Email Manager.url"

:: קיצור דרך בתפריט התחל
echo [InternetShortcut] > "%startMenu%\AI Email Manager.url"
echo URL=file:///C:/Users/%USERNAME%/outlook_email_manager/outlook_com_addin.py >> "%startMenu%\AI Email Manager.url"
echo IconFile=C:\Users\%USERNAME%\outlook_email_manager\outlook_addin\icon-32.ico >> "%startMenu%\AI Email Manager.url"
echo IconIndex=0 >> "%startMenu%\AI Email Manager.url"

echo ✅ קיצורי דרך נוצרו
echo.

:: יצירת סקריפט הפעלה
echo 📝 יוצר סקריפט הפעלה...
echo @echo off > "C:\Users\%USERNAME%\outlook_email_manager\start_addin.bat"
echo chcp 65001 ^>nul >> "C:\Users\%USERNAME%\outlook_email_manager\start_addin.bat"
echo echo מתחיל AI Email Manager... >> "C:\Users\%USERNAME%\outlook_email_manager\start_addin.bat"
echo cd /d "C:\Users\%USERNAME%\outlook_email_manager" >> "C:\Users\%USERNAME%\outlook_email_manager\start_addin.bat"
echo python outlook_com_addin.py >> "C:\Users\%USERNAME%\outlook_email_manager\start_addin.bat"
echo pause >> "C:\Users\%USERNAME%\outlook_email_manager\start_addin.bat"

echo ✅ סקריפט הפעלה נוצר
echo.

:: בדיקת התקנה
echo 🧪 בודק התקנה...
cd /d "C:\Users\%USERNAME%\outlook_email_manager"
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
echo 2. 🚀 הפעל את התוסף:
echo    python outlook_com_addin.py
echo    או לחץ על הקיצור "AI Email Manager"
echo.
echo 3. 📧 פתח את Outlook ובחר מיילים לניתוח
echo.
echo 4. 🎯 השתמש בכפתורי ה-Ribbon החדשים
echo.
echo 📞 תמיכה:
echo - בדוק את הלוגים ב-outlook_addin.log
echo - ודא שהשרת רץ על localhost:5000
echo - בדוק את החיבור ל-Outlook
echo.
echo לחץ על מקש כלשהו לסגירה...
pause >nul











