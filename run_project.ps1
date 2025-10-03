# Outlook Email Manager - Enhanced Run Script
# Parameters
param(
    [switch]$StopOnly,
    [int]$Port = 5000
)
# סקריפט הרצה משופר למערכת ניהול מיילים חכמה

# Save this file with UTF-8 with BOM encoding to ensure emojis display correctly in PowerShell.

# הגדרת קידוד UTF-8 לתמיכה בעברית
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "Outlook Email Manager with AI - Main Run Script" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Navigate to project directory
Set-Location $PSScriptRoot
Write-Host "Working directory: $PWD" -ForegroundColor Yellow
Write-Host ""

# Helper: Stop existing servers robustly
function Stop-ExistingServers {
    param([int]$ListenPort)

    Write-Host "Stopping existing servers..." -ForegroundColor Red
    $killed = 0

    # 1) Kill by port (IPv4/IPv6)
    try {
        $conns = Get-NetTCPConnection -LocalPort $ListenPort -ErrorAction SilentlyContinue
        foreach ($c in $conns) {
            try { Stop-Process -Id $c.OwningProcess -Force -ErrorAction SilentlyContinue; $killed++ } catch {}
        }
    } catch {}

    # Fallback using netstat (older PS)
    try {
        $lines = (& netstat -ano -p tcp) 2>$null | Select-String "\s$ListenPort\s"
        foreach ($ln in $lines) {
            $parts = ($ln -split "\s+") | Where-Object { $_ -ne '' }
            $procId = $parts[-1]
            if ($procId -as [int]) { try { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue; $killed++ } catch {} }
        }
    } catch {}

    # 2) Kill python processes started from this workspace
    $workspace = (Get-Location).Path
    try {
        $procs = Get-CimInstance Win32_Process | Where-Object { $_.Name -match 'python' -and $_.CommandLine -match [regex]::Escape($workspace) }
        foreach ($p in $procs) {
            try { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue; $killed++ } catch {}
        }
    } catch {}

    # 3) Common dev servers
    try { Get-Process -Name flask,gunicorn,pythonw -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue; $killed++ } } catch {}

    Start-Sleep -Seconds 1
    # ניסיון אחרון לשחרר את הפורט ע"י netstat
    for ($i=0; $i -lt 5; $i++) {
        $still = Get-NetTCPConnection -LocalPort $ListenPort -ErrorAction SilentlyContinue
        if ($still -eq $null) { break }
        $lines = (& netstat -ano -p tcp) 2>$null | Select-String "\s$ListenPort\s"
        foreach ($ln in $lines) {
            $parts = ($ln -split "\s+") | Where-Object { $_ -ne '' }
            $procId = $parts[-1]
            if ($procId -as [int]) { try { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue; $killed++ } catch {} }
        }
        Start-Sleep -Milliseconds 300
    }
    if ($killed -gt 0) { Write-Host "Stopped $killed existing server process(es)." -ForegroundColor Yellow }
    else { Write-Host "No existing servers found." -ForegroundColor Green }
}

# Helper: Ensure port is free
function Test-PortFree {
    param([int]$ListenPort)
    $busy = Get-NetTCPConnection -LocalPort $ListenPort -ErrorAction SilentlyContinue
    return ($busy -eq $null)
}

# Check Python
Write-Host "Checking Python installation..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    if ($pythonVersion -match "Python 3\.[0-9]+|Python [4-9]") {
        Write-Host "Python found: $pythonVersion" -ForegroundColor Green
    } else {
        Write-Host "Python 3.6+ required! Found: $pythonVersion" -ForegroundColor Red
        Write-Host "Please install Python 3.6 or higher from https://python.org" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
} catch {
    Write-Host "Python not found! Please install Python 3.6+ from https://python.org" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check pip
Write-Host "Checking pip..." -ForegroundColor Yellow
try {
    $pipVersion = pip --version 2>&1
    if ($pipVersion) { Write-Host "pip is OK" -ForegroundColor Green } else { throw }
} catch {
    Write-Host "pip not found! Please install pip" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

Write-Host ""

Stop-ExistingServers -ListenPort $Port

# אם קיים דגל אתחול – ננקה אותו אם נשאר מריצה קודמת (ללא בלוק if)
$restartFlag = Join-Path $PSScriptRoot 'restart.flag'
try { if (Test-Path $restartFlag) { Remove-Item $restartFlag -Force -ErrorAction SilentlyContinue } } catch {}

if ($StopOnly) { Write-Host "StopOnly mode: done." -ForegroundColor Yellow; exit 0 }

# Optional: clear Python caches for a clean start
try { Remove-Item -Force -Recurse -ErrorAction SilentlyContinue "__pycache__"; Get-ChildItem -Recurse -Include *.pyc -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue } catch {}

Write-Host ""

# בדיקות – אין קובץ בדיקות, מדלגים
Write-Host "Skipping tests (none)" -ForegroundColor Yellow
Write-Host ""

# התקנת תלויות
Write-Host "Installing dependencies..." -ForegroundColor Yellow
try {
    pip install -r requirements.txt --quiet
    Write-Host "Dependencies installed successfully!" -ForegroundColor Green
} catch {
    Write-Host "Error installing dependencies!" -ForegroundColor Red
    Write-Host "Trying to install manually..." -ForegroundColor Yellow
    pip install flask==2.3.3 pywin32 google-generativeai==0.3.2
}

Write-Host ""

# בדיקת קבצים נדרשים
Write-Host "Checking required files..." -ForegroundColor Yellow
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
        Write-Host "$file" -ForegroundColor Green
    } else {
        Write-Host "$file - MISSING!" -ForegroundColor Red
        $missingFiles += $file
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host ""
    Write-Host "Missing required files! Cannot start server." -ForegroundColor Red
    Write-Host "Missing files:" -ForegroundColor Red
    $missingFiles | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""

# בדיקת Outlook
Write-Host "Checking Outlook status..." -ForegroundColor Yellow
try {
    $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    if ($outlookProcess) {
        Write-Host "Outlook is running" -ForegroundColor Green
    } else {
        Write-Host "Outlook is not running - system will use sample data" -ForegroundColor Yellow
        Write-Host "   To use real emails, please start Outlook first" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Could not check Outlook status" -ForegroundColor Yellow
}

Write-Host ""

# בדיקת API Key
Write-Host "Checking AI configuration..." -ForegroundColor Yellow
try {
    $configContent = Get-Content "config.py" -Raw
    if ($configContent -match "your_api_key_here") {
        Write-Host "Gemini API Key not configured - AI features will be limited" -ForegroundColor Yellow
        Write-Host "   To enable full AI features, update config.py with your Gemini API key" -ForegroundColor Yellow
    } else {
        Write-Host "AI configuration looks good" -ForegroundColor Green
    }
} catch {
    Write-Host "Could not check AI configuration" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Starting Outlook Email Manager with AI..." -ForegroundColor Yellow
Write-Host "Server will be available at: " -ForegroundColor Green -NoNewline
Write-Host "http://localhost:5000" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press Ctrl+C to stop the server" -ForegroundColor Red

# הרצת השרת בלולאה: אם נדרש אתחול – נריץ שוב בלי לצאת מהטרמינל
while ($true) {
    $relaunch = $false
    try {
    # סינון לוגים מיותרים של gRPC/ABSL/Gemini: הגדרת משתני סביבה לפני הרצת פייתון
    $env:TF_CPP_MIN_LOG_LEVEL = '3'
    $env:GRPC_VERBOSITY = 'NONE'
    $env:GLOG_minloglevel = '4'
    $env:GRPC_TRACE = ''
    $env:ABSL_LOG_LEVEL = 'FATAL'
    $env:TERMINAL_LOG_LEVEL = 'CRITICAL'

    # ניתוב פלט השרת לקבצי לוג כדי לשמור את הטרמינל נקי מהודעות ספריות חיצוניות
    $logDir = Join-Path $PSScriptRoot 'logs'
    try { New-Item -ItemType Directory -Path $logDir -Force | Out-Null } catch {}
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outLog = Join-Path $logDir "server-out-$timestamp.log"
    $errLog = Join-Path $logDir "server-err-$timestamp.log"

    $args = @('app_with_ai.py')
    $server = Start-Process -FilePath python -ArgumentList $args -WorkingDirectory $PSScriptRoot -PassThru -NoNewWindow -RedirectStandardOutput $outLog -RedirectStandardError $errLog
    Write-Host "Server PID: $($server.Id)" -ForegroundColor DarkCyan
    Write-Host "(server output redirected to $outLog, errors to $errLog)" -ForegroundColor DarkGray
        try {
            Wait-Process -Id $server.Id
    } finally {
        Write-Host "Stopping server process..." -ForegroundColor Yellow
        try { if (-not $server.HasExited) { Stop-Process -Id $server.Id -Force -ErrorAction SilentlyContinue } } catch {}
        Stop-ExistingServers -ListenPort $Port

        # אם נוצר דגל במהלך הריצה – נסמן להפעלה מחדש
        if (Test-Path $restartFlag) {
            try { Remove-Item $restartFlag -Force -ErrorAction SilentlyContinue } catch {}
            $relaunch = $true
        }

        # בדיקת קוד יציאה מיוחד (222) – מעיד על אתחול
        try { $exitCode = $server.ExitCode } catch { $exitCode = $null }
        if ($exitCode -eq 222) { $relaunch = $true }
    }
    } catch {
        # אם השרת יצא עם קוד 222 (אתחול), נטפל בו כלוגיקה של אתחול
        if ($LASTEXITCODE -eq 222 -or (Test-Path $restartFlag)) {
            try { Remove-Item $restartFlag -Force -ErrorAction SilentlyContinue } catch {}
            $relaunch = $true
        }
        Write-Host ""
        Write-Host "Error starting server!" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "Troubleshooting tips:" -ForegroundColor Yellow
        Write-Host "1. Make sure all dependencies are installed" -ForegroundColor White
        Write-Host "2. Check if port $Port is available" -ForegroundColor White
        Write-Host "3. Try running: python app_with_ai.py" -ForegroundColor White
        Write-Host "4. Check the error messages above" -ForegroundColor White
    }

    # תמיד ננסה להרים מחדש (Ctrl+C יפסיק את הסקריפט)
    Write-Host "Relaunching server..." -ForegroundColor Yellow
    Start-Sleep -Seconds 1
    continue
}

Write-Host ""
Write-Host "Server stopped. Thank you for using Outlook Email Manager!" -ForegroundColor Green
Read-Host "Press Enter to close"
