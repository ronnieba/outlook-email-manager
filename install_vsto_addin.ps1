# Install VSTO Add-in Script
# הפעל כמנהל מערכת!

Write-Host "מתקין VSTO Add-in..." -ForegroundColor Green

# נתיב ל-VSTO
$vstoPath = "$PSScriptRoot\AIEmailManagerAddin\bin\Debug\AIEmailManagerAddin.vsto"
$manifestPath = "$PSScriptRoot\AIEmailManagerAddin\bin\Debug\AIEmailManagerAddin.dll.manifest"

Write-Host "נתיב VSTO: $vstoPath"

# בדיקה שהקובץ קיים
if (-Not (Test-Path $vstoPath)) {
    Write-Host "שגיאה: הקובץ VSTO לא נמצא!" -ForegroundColor Red
    Write-Host "ודא שבנית את הפרויקט ב-Debug mode" -ForegroundColor Yellow
    exit 1
}

# הוספת Trust ל-VSTO
Write-Host "מוסיף אמון (Trust) ל-VSTO..."

# יצירת Registry Key לתוסף
$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\AIEmailManagerAddin"

# בדיקה אם כבר קיים
if (Test-Path $regPath) {
    Write-Host "מוחק רישום קיים..."
    Remove-Item -Path $regPath -Recurse -Force
}

# יצירת Key חדש
New-Item -Path $regPath -Force | Out-Null

# הוספת ערכים
Set-ItemProperty -Path $regPath -Name "Description" -Value "AI Email Manager for Outlook" -Type String
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "AI Email Manager" -Type String
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $regPath -Name "Manifest" -Value "$vstoPath|vstolocal" -Type String

Write-Host "✓ התוסף נרשם בהצלחה!" -ForegroundColor Green
Write-Host ""
Write-Host "עכשיו:" -ForegroundColor Yellow
Write-Host "1. פתח את Outlook" -ForegroundColor White
Write-Host "2. File → Options → Add-ins" -ForegroundColor White
Write-Host "3. בחר 'COM Add-ins' → Go" -ForegroundColor White
Write-Host "4. סמן את 'AI Email Manager'" -ForegroundColor White
Write-Host ""
Write-Host "אם עדיין לא עובד, הרץ Visual Studio כמנהל ובנה מחדש!" -ForegroundColor Cyan

Read-Host "לחץ Enter לסגירה"
