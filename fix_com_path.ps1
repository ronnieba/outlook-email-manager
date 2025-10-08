# תיקון נתיב התוסף COM
Write-Host "========================================"
Write-Host "תיקון נתיב התוסף COM"
Write-Host "========================================"

Write-Host ""
Write-Host "1. בודק את הנתיב הנוכחי..."
$currentPath = Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}\LocalServer32" -Name "(Default)" -ErrorAction SilentlyContinue
if ($currentPath) {
    Write-Host "נתיב נוכחי: $($currentPath.'(Default)')"
} else {
    Write-Host "לא נמצא נתיב נוכחי"
}

Write-Host ""
Write-Host "2. מעדכן את הנתיב..."
$newPath = "C:\Users\ronni\outlook_email_manager\dist\exe_addin.exe"
Set-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}\LocalServer32" -Name "(Default)" -Value $newPath

Write-Host ""
Write-Host "3. בודק את הנתיב החדש..."
$updatedPath = Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}\LocalServer32" -Name "(Default)"
Write-Host "נתיב חדש: $($updatedPath.'(Default)')"

Write-Host ""
Write-Host "4. בודק אם הקובץ קיים..."
if (Test-Path $newPath) {
    Write-Host "✅ הקובץ קיים: $newPath"
} else {
    Write-Host "❌ הקובץ לא קיים: $newPath"
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


