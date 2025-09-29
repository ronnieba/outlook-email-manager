# הגדרת קידוד UTF-8 לתמיכה בעברית
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "Starting Outlook Email Manager..." -ForegroundColor Green
Write-Host ""

Set-Location $PSScriptRoot

Write-Host "Stopping any existing servers..." -ForegroundColor Red
# עצירת תהליכי Python שרצים על פורט 5000
$processes = Get-Process -Name "python" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -like "*app_outlook_fixed.py*" }
if ($processes) {
    Write-Host "Found existing server processes, stopping them..." -ForegroundColor Yellow
    $processes | Stop-Process -Force
    Start-Sleep -Seconds 2
    Write-Host "Existing servers stopped." -ForegroundColor Green
} else {
    Write-Host "No existing servers found." -ForegroundColor Green
}

Write-Host ""
Write-Host "Checking dependencies..." -ForegroundColor Yellow
pip install -r requirements.txt

Write-Host ""
Write-Host "Starting server..." -ForegroundColor Cyan
python app_outlook_fixed.py

Read-Host "Press Enter to close"


