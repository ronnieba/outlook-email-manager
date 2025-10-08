@echo off
echo ========================================
echo תיקון רישום PythonCOM
echo ========================================

echo.
echo 1. בודק את הרישום הנוכחי...
reg query "HKEY_CLASSES_ROOT\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"

echo.
echo 2. מוסיף רישום PythonCOM...
reg add "HKEY_CLASSES_ROOT\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}\PythonCOM" /ve /t REG_SZ /d "inproc_com_addin.InprocCOMAddin" /f

echo.
echo 3. מוסיף רישום PythonCOMPath...
reg add "HKEY_CLASSES_ROOT\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}\PythonCOMPath" /ve /t REG_SZ /d "C:\Users\ronni\outlook_email_manager" /f

echo.
echo 4. בודק את הרישום החדש...
reg query "HKEY_CLASSES_ROOT\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}\PythonCOM"
reg query "HKEY_CLASSES_ROOT\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}\PythonCOMPath"

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


