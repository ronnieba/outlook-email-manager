# Quick Start Script for Outlook Email Manager
# סקריפט הרצה מהיר למערכת ניהול מיילים חכמה

# ניקוי המסך לפני הרצה
Clear-Host

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "🚀 Quick Start - Outlook Email Manager" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# מעבר לספריית הפרויקט
Set-Location $PSScriptRoot
Write-Host "Working directory: $PWD" -ForegroundColor Yellow
Write-Host ""

# עצירת שרתים קיימים
Write-Host "🛑 Stopping existing servers..." -ForegroundColor Red
try {
    $processes = Get-Process -Name "python" -ErrorAction SilentlyContinue | Where-Object { 
        $_.CommandLine -like "*app_with_ai.py*" -or 
        $_.CommandLine -like "*app_outlook_fixed.py*" -or
        $_.CommandLine -like "*app.py*"
    }
    if ($processes) {
        Write-Host "Found existing server processes, stopping them..." -ForegroundColor Yellow
        $processes | Stop-Process -Force
        Start-Sleep -Seconds 2
        Write-Host "✅ Existing servers stopped." -ForegroundColor Green
    } else {
        Write-Host "✅ No existing servers found." -ForegroundColor Green
    }
} catch {
    Write-Host "⚠️ Could not check for existing processes" -ForegroundColor Yellow
}

Write-Host ""

# בדיקת Python
Write-Host "🐍 Checking Python installation..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    if ($pythonVersion -match "Python 3\.[6-9]|Python 3\.[1-9][0-9]|Python [4-9]") {
        Write-Host "✅ Python found: $pythonVersion" -ForegroundColor Green
    } else {
        Write-Host "❌ Python 3.6+ required! Found: $pythonVersion" -ForegroundColor Red
        Write-Host "Please install Python 3.6 or higher from https://python.org" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
} catch {
    Write-Host "❌ Python not found! Please install Python 3.6+ from https://python.org" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# בדיקת קבצים נדרשים
Write-Host "📋 Checking required files..." -ForegroundColor Yellow
$requiredFiles = @(
    "app_with_ai.py",
    "ai_analyzer.py", 
    "config.py",
    "user_profile_manager.py",
    "templates\index.html",
    "requirements.txt"
)

$missingFiles = @()
foreach ($file in $requiredFiles) {
    if (Test-Path $file) {
        Write-Host "✅ $file" -ForegroundColor Green
    } else {
        Write-Host "❌ $file - MISSING!" -ForegroundColor Red
        $missingFiles += $file
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host ""
    Write-Host "❌ Missing required files! Cannot start server." -ForegroundColor Red
    Write-Host "Missing files:" -ForegroundColor Red
    $missingFiles | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# התקנת תלויות
Write-Host "📦 Installing dependencies..." -ForegroundColor Yellow
try {
    pip install -r requirements.txt --quiet
    Write-Host "✅ Dependencies installed successfully!" -ForegroundColor Green
} catch {
    Write-Host "❌ Error installing dependencies!" -ForegroundColor Red
    Write-Host "Trying to install manually..." -ForegroundColor Yellow
    pip install flask==2.3.3 pywin32 google-generativeai==0.3.2
}

Write-Host ""

# בדיקת Outlook
Write-Host "📧 Checking Outlook status..." -ForegroundColor Yellow
try {
    $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlookProcess) {
        Write-Host "✅ Outlook is running" -ForegroundColor Green
    } else {
        Write-Host "⚠️ Outlook is not running - system will use sample data" -ForegroundColor Yellow
        Write-Host "   To use real emails, please start Outlook first" -ForegroundColor Yellow
    }
} catch {
    Write-Host "⚠️ Could not check Outlook status" -ForegroundColor Yellow
}

Write-Host ""

# בדיקת API Key
Write-Host "🤖 Checking AI configuration..." -ForegroundColor Yellow
try {
    $configContent = Get-Content "config.py" -Raw
    if ($configContent -match "your_api_key_here") {
        Write-Host "⚠️ Gemini API Key not configured - AI features will be limited" -ForegroundColor Yellow
        Write-Host "   To enable full AI features, update config.py with your Gemini API key" -ForegroundColor Yellow
    } else {
        Write-Host "✅ AI configuration looks good" -ForegroundColor Green
    }
} catch {
    Write-Host "⚠️ Could not check AI configuration" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "🚀 Starting Outlook Email Manager with AI..." -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "🌐 Server will be available at: http://localhost:5000" -ForegroundColor Cyan
Write-Host "🛑 Press Ctrl+C to stop the server" -ForegroundColor Red
Write-Host ""

# הרצת השרת
try {
    python app_with_ai.py
} catch {
    Write-Host ""
    Write-Host "❌ Error starting server!" -ForegroundColor Red
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting tips:" -ForegroundColor Yellow
    Write-Host "1. Make sure all dependencies are installed" -ForegroundColor White
    Write-Host "2. Check if port 5000 is available" -ForegroundColor White
    Write-Host "3. Try running: python app_with_ai.py" -ForegroundColor White
    Write-Host "4. Check the error messages above" -ForegroundColor White
}

Write-Host ""
Write-Host "👋 Server stopped. Thank you for using Outlook Email Manager!" -ForegroundColor Green
Read-Host "Press Enter to close"
