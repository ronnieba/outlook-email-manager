@echo off
echo ========================================
echo תיקון נתיב התוסף COM
echo ========================================

echo.
echo 1. בודק את הנתיב הנוכחי...
reg query "HKEY_CLASSES_ROOT\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}\LocalServer32"

echo.
echo 2. מעדכן את הנתיב...
reg add "HKEY_CLASSES_ROOT\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}\LocalServer32" /ve /t REG_SZ /d "C:\Users\ronni\outlook_email_manager\dist\exe_addin.exe" /f

echo.
echo 3. בודק את הנתיב החדש...
reg query "HKEY_CLASSES_ROOT\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}\LocalServer32"

echo.
echo 4. בודק אם הקובץ קיים...
if exist "C:\Users\ronni\outlook_email_manager\dist\exe_addin.exe" (
    echo ✅ הקובץ קיים: C:\Users\ronni\outlook_email_manager\dist\exe_addin.exe
) else (
    echo ❌ הקובץ לא קיים: C:\Users\ronni\outlook_email_manager\dist\exe_addin.exe
)

echo.
echo ========================================
echo התיקון הושלם!
echo ========================================
echo.
echo עכשיו:
echo 1. סגור את Outlook לחלוטין
echo 2. פתח את Outlook מחדש
echo 3. בדוק אם התוסף נטען
echo.
pause


