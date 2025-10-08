# תיקון רישום PythonCOM
Write-Host "========================================"
Write-Host "תיקון רישום PythonCOM"
Write-Host "========================================"

Write-Host ""
Write-Host "1. בודק את הרישום הנוכחי..."
$currentReg = Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" -ErrorAction SilentlyContinue
if ($currentReg) {
    Write-Host "CLSID קיים"
} else {
    Write-Host "CLSID לא קיים"
}

Write-Host ""
Write-Host "2. מוסיף רישום PythonCOM..."
Set-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" -Name "PythonCOM" -Value "inproc_com_addin.InprocCOMAddin"

Write-Host ""
Write-Host "3. מוסיף רישום PythonCOMPath..."
Set-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" -Name "PythonCOMPath" -Value "C:\Users\ronni\outlook_email_manager"

Write-Host ""
Write-Host "4. בודק את הרישום החדש..."
$pythonCOM = Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" -Name "PythonCOM" -ErrorAction SilentlyContinue
if ($pythonCOM) {
    Write-Host "PythonCOM: $($pythonCOM.PythonCOM)"
} else {
    Write-Host "PythonCOM לא נמצא"
}

$pythonCOMPath = Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" -Name "PythonCOMPath" -ErrorAction SilentlyContinue
if ($pythonCOMPath) {
    Write-Host "PythonCOMPath: $($pythonCOMPath.PythonCOMPath)"
} else {
    Write-Host "PythonCOMPath לא נמצא"
}

Write-Host ""
Write-Host "========================================"
Write-Host "התיקון הושלם!"
Write-Host "========================================"
Write-Host ""
Write-Host "עכשיו:"
Write-Host "1. סגור את Outlook לחלוטין"
Write-Host "2. פתח את Outlook מחדש"
Write-Host "3. בדוק אם התוסף נטען"
Write-Host ""
Read-Host "לחץ Enter להמשך"


