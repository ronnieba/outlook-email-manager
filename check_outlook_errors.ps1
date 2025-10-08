# בדיקת שגיאות Outlook ב-Event Viewer

Write-Host "מחפש שגיאות Outlook..." -ForegroundColor Yellow
Write-Host ""

# קבלת 10 האירועים האחרונים של Application Errors
$errors = Get-EventLog -LogName Application -Source "Outlook" -EntryType Error -Newest 10 -ErrorAction SilentlyContinue

if ($errors) {
    Write-Host "נמצאו שגיאות Outlook:" -ForegroundColor Red
    Write-Host "===================="
    foreach ($error in $errors) {
        Write-Host ""
        Write-Host "זמן: $($error.TimeGenerated)" -ForegroundColor Cyan
        Write-Host "הודעה: $($error.Message)" -ForegroundColor White
        Write-Host "--------------------"
    }
} else {
    Write-Host "לא נמצאו שגיאות Outlook ב-Event Log" -ForegroundColor Green
}

Write-Host ""
Write-Host "מחפש שגיאות VSTO..." -ForegroundColor Yellow

# חיפוש שגיאות VSTO
$vstoErrors = Get-EventLog -LogName Application -Source "VSTO*" -EntryType Error -Newest 10 -ErrorAction SilentlyContinue

if ($vstoErrors) {
    Write-Host "נמצאו שגיאות VSTO:" -ForegroundColor Red
    Write-Host "===================="
    foreach ($error in $vstoErrors) {
        Write-Host ""
        Write-Host "זמן: $($error.TimeGenerated)" -ForegroundColor Cyan
        Write-Host "הודעה: $($error.Message)" -ForegroundColor White
        Write-Host "--------------------"
    }
} else {
    Write-Host "לא נמצאו שגיאות VSTO ב-Event Log" -ForegroundColor Green
}

Read-Host "`nלחץ Enter לסגירה"
