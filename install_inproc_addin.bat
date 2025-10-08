@echo off
echo ========================================
echo התקנת תוסף COM כ-InprocServer32
echo ========================================

echo.
echo 1. מבטל רישום תוספים קודמים...
reg delete "HKEY_CLASSES_ROOT\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" /f 2>nul
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\InprocCOMAddin.Addin" /f 2>nul

echo.
echo 2. רושם את התוסף החדש...
python inproc_com_addin.py

echo.
echo 3. מעדכן את הרישום ל-InprocServer32...
reg add "HKEY_CLASSES_ROOT\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}\InprocServer32" /ve /t REG_SZ /d "C:\Users\ronni\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.13_qbz5n2kfra8p0\LocalCache\local-packages\Python313\site-packages\pywin32_system32\pythoncom313.dll" /f
reg add "HKEY_CLASSES_ROOT\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}\InprocServer32" /v "ThreadingModel" /t REG_SZ /d "Apartment" /f

echo.
echo 4. מוחק את LocalServer32 אם קיים...
reg delete "HKEY_CLASSES_ROOT\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}\LocalServer32" /f 2>nul

echo.
echo 5. מעדכן את רישום Outlook...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\InprocCOMAddin.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\InprocCOMAddin.Addin" /v "FriendlyName" /t REG_SZ /d "Inproc COM Addin" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\InprocCOMAddin.Addin" /v "Description" /t REG_SZ /d "Inproc COM Addin for Testing" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\InprocCOMAddin.Addin" /v "Authentication" /t REG_DWORD /d 1 /f

echo.
echo 6. בודק את הרישום...
reg query "HKEY_CLASSES_ROOT\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"
echo.
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\InprocCOMAddin.Addin"

echo.
echo ========================================
echo ההתקנה הושלמה!
echo ========================================
echo.
echo עכשיו:
echo 1. סגור את Outlook לחלוטין
echo 2. פתח את Outlook מחדש
echo 3. בדוק אם התוסף נטען
echo 4. בדוק אם נוצרו קבצי טסט בתיקיית TEMP
echo.
pause


