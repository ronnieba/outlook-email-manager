# AI Email Manager - התקנת תוסף COM מתקדמת
# PowerShell Script for Advanced COM Add-in Installation

param(
    [switch]$Force,
    [switch]$Silent,
    [string]$InstallPath = "$env:USERPROFILE\outlook_email_manager"
)

# הגדרת קידוד UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.OutputEncoding = [System.Text.Encoding]::UTF8

# פונקציות עזר
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    if (-not $Silent) {
        Write-Host $Message -ForegroundColor $Color
    }
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-PythonInstalled {
    try {
        $pythonVersion = python --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            return $true
        }
    } catch {
        return $false
    }
    return $false
}

function Test-OutlookInstalled {
    try {
        $officeKeys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Office" -ErrorAction SilentlyContinue
        return $officeKeys.Count -gt 0
    } catch {
        return $false
    }
}

function Install-PythonDependencies {
    Write-ColorOutput "📦 מתקין תלויות Python..." "Yellow"
    
    $packages = @(
        "flask==2.3.3",
        "flask-cors==4.0.0", 
        "pywin32>=307",
        "google-generativeai==0.3.2",
        "requests",
        "pythoncom"
    )
    
    foreach ($package in $packages) {
        try {
            Write-ColorOutput "  מתקין $package..." "Gray"
            pip install $package --quiet
            if ($LASTEXITCODE -eq 0) {
                Write-ColorOutput "  ✅ $package הותקן" "Green"
            } else {
                Write-ColorOutput "  ❌ שגיאה בהתקנת $package" "Red"
                return $false
            }
        } catch {
            Write-ColorOutput "  ❌ שגיאה בהתקנת $package: $($_.Exception.Message)" "Red"
            return $false
        }
    }
    
    return $true
}

function Copy-AddinFiles {
    param([string]$SourcePath, [string]$DestinationPath)
    
    Write-ColorOutput "📋 מעתיק קבצי התוסף..." "Yellow"
    
    $filesToCopy = @(
        "outlook_com_addin.py",
        "outlook_addin\manifest.xml",
        "outlook_addin\taskpane.html", 
        "outlook_addin\taskpane.js",
        "outlook_addin\taskpane.css",
        "outlook_addin\icon-32.ico",
        "outlook_addin\icon-64.ico"
    )
    
    foreach ($file in $filesToCopy) {
        $sourceFile = Join-Path $SourcePath $file
        $destFile = Join-Path $DestinationPath $file
        
        if (Test-Path $sourceFile) {
            try {
                $destDir = Split-Path $destFile -Parent
                if (-not (Test-Path $destDir)) {
                    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                }
                
                Copy-Item $sourceFile $destFile -Force
                Write-ColorOutput "  ✅ הועתק: $file" "Green"
            } catch {
                Write-ColorOutput "  ❌ שגיאה בהעתקת $file: $($_.Exception.Message)" "Red"
                return $false
            }
        } else {
            Write-ColorOutput "  ⚠️ קובץ לא נמצא: $file" "Yellow"
        }
    }
    
    return $true
}

function Register-COMAddin {
    Write-ColorOutput "🔧 רושם תוסף COM ב-Windows Registry..." "Yellow"
    
    try {
        # רישום התוסף
        $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin"
        New-Item -Path $regPath -Force | Out-Null
        
        Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord
        Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "AI Email Manager" -Type String
        Set-ItemProperty -Path $regPath -Name "Description" -Value "ניתוח חכם של מיילים ופגישות עם AI" -Type String
        Set-ItemProperty -Path $regPath -Name "Manifest" -Value "$InstallPath\outlook_addin\manifest.xml" -Type String
        
        # הגדרות התוסף
        $settingsPath = "HKCU:\Software\AIEmailManager"
        New-Item -Path $settingsPath -Force | Out-Null
        
        Set-ItemProperty -Path $settingsPath -Name "Version" -Value "1.0.0" -Type String
        Set-ItemProperty -Path $settingsPath -Name "InstallPath" -Value $InstallPath -Type String
        Set-ItemProperty -Path $settingsPath -Name "ServerURL" -Value "http://localhost:5000" -Type String
        Set-ItemProperty -Path $settingsPath -Name "AutoAnalyze" -Value 1 -Type DWord
        Set-ItemProperty -Path $settingsPath -Name "AnalyzeMeetings" -Value 1 -Type DWord
        Set-ItemProperty -Path $settingsPath -Name "ShowNotifications" -Value 1 -Type DWord
        
        Write-ColorOutput "✅ תוסף נרשם ב-Windows Registry" "Green"
        return $true
        
    } catch {
        Write-ColorOutput "❌ שגיאה ברישום התוסף: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Create-Shortcuts {
    Write-ColorOutput "🔗 יוצר קיצורי דרך..." "Yellow"
    
    try {
        $desktop = [Environment]::GetFolderPath("Desktop")
        $startMenu = [Environment]::GetFolderPath("StartMenu")
        
        # קיצור דרך על שולחן העבודה
        $desktopShortcut = Join-Path $desktop "AI Email Manager.lnk"
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($desktopShortcut)
        $Shortcut.TargetPath = "python.exe"
        $Shortcut.Arguments = "`"$InstallPath\outlook_com_addin.py`""
        $Shortcut.WorkingDirectory = $InstallPath
        $Shortcut.Description = "AI Email Manager - תוסף ניתוח מיילים חכם"
        $Shortcut.Save()
        
        # קיצור דרך בתפריט התחל
        $startShortcut = Join-Path $startMenu "AI Email Manager.lnk"
        $Shortcut2 = $WshShell.CreateShortcut($startShortcut)
        $Shortcut2.TargetPath = "python.exe"
        $Shortcut2.Arguments = "`"$InstallPath\outlook_com_addin.py`""
        $Shortcut2.WorkingDirectory = $InstallPath
        $Shortcut2.Description = "AI Email Manager - תוסף ניתוח מיילים חכם"
        $Shortcut2.Save()
        
        Write-ColorOutput "✅ קיצורי דרך נוצרו" "Green"
        return $true
        
    } catch {
        Write-ColorOutput "❌ שגיאה ביצירת קיצורי דרך: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Create-StartupScript {
    Write-ColorOutput "📝 יוצר סקריפט הפעלה..." "Yellow"
    
    try {
        $startupScript = Join-Path $InstallPath "start_addin.bat"
        $scriptContent = @"
@echo off
chcp 65001 >nul
title AI Email Manager
echo.
echo ========================================
echo    AI Email Manager - תוסף Outlook
echo ========================================
echo.
echo מתחיל תוסף...
cd /d "$InstallPath"
python outlook_com_addin.py
echo.
echo לחץ על מקש כלשהו לסגירה...
pause >nul
"@
        
        Set-Content -Path $startupScript -Value $scriptContent -Encoding UTF8
        Write-ColorOutput "✅ סקריפט הפעלה נוצר" "Green"
        return $true
        
    } catch {
        Write-ColorOutput "❌ שגיאה ביצירת סקריפט הפעלה: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Test-Installation {
    Write-ColorOutput "🧪 בודק התקנה..." "Yellow"
    
    try {
        # בדיקת COM
        $comTest = python -c "import win32com.client; print('COM: OK')" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-ColorOutput "✅ בדיקת COM עברה בהצלחה" "Green"
        } else {
            Write-ColorOutput "❌ שגיאה בבדיקת COM" "Red"
            return $false
        }
        
        # בדיקת קבצים
        $requiredFiles = @(
            "outlook_com_addin.py",
            "outlook_addin\manifest.xml"
        )
        
        foreach ($file in $requiredFiles) {
            $filePath = Join-Path $InstallPath $file
            if (Test-Path $filePath) {
                Write-ColorOutput "✅ קובץ קיים: $file" "Green"
            } else {
                Write-ColorOutput "❌ קובץ חסר: $file" "Red"
                return $false
            }
        }
        
        return $true
        
    } catch {
        Write-ColorOutput "❌ שגיאה בבדיקת התקנה: $($_.Exception.Message)" "Red"
        return $false
    }
}

# התחלת התקנה
Write-ColorOutput ""
Write-ColorOutput "========================================" "Cyan"
Write-ColorOutput "   AI Email Manager - התקנת תוסף COM" "Cyan"
Write-ColorOutput "========================================" "Cyan"
Write-ColorOutput ""

# בדיקת הרשאות מנהל
if (-not (Test-Administrator)) {
    Write-ColorOutput "❌ נדרשות הרשאות מנהל להתקנה" "Red"
    Write-ColorOutput "הפעל את הסקריפט כמנהל (Run as Administrator)" "Yellow"
    if (-not $Silent) {
        Read-Host "לחץ Enter לסגירה"
    }
    exit 1
}

Write-ColorOutput "✅ הרשאות מנהל מאושרות" "Green"

# בדיקת Python
Write-ColorOutput "🔍 בודק Python..." "Yellow"
if (-not (Test-PythonInstalled)) {
    Write-ColorOutput "❌ Python לא מותקן או לא נמצא ב-PATH" "Red"
    Write-ColorOutput "אנא התקן Python 3.8+ מ-https://www.python.org/downloads/" "Yellow"
    if (-not $Silent) {
        Read-Host "לחץ Enter לסגירה"
    }
    exit 1
}

Write-ColorOutput "✅ Python מותקן" "Green"

# בדיקת Outlook
Write-ColorOutput "🔍 בודק Microsoft Outlook..." "Yellow"
if (-not (Test-OutlookInstalled)) {
    Write-ColorOutput "❌ Microsoft Outlook לא מותקן" "Red"
    Write-ColorOutput "אנא התקן Microsoft Outlook 2016+ לפני המשך" "Yellow"
    if (-not $Silent) {
        Read-Host "לחץ Enter לסגירה"
    }
    exit 1
}

Write-ColorOutput "✅ Microsoft Outlook מותקן" "Green"

# יצירת תיקיית התקנה
Write-ColorOutput "📁 יוצר תיקיית התקנה..." "Yellow"
try {
    if (-not (Test-Path $InstallPath)) {
        New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
    }
    
    $subdirs = @("outlook_addin", "logs", "templates")
    foreach ($subdir in $subdirs) {
        $subdirPath = Join-Path $InstallPath $subdir
        if (-not (Test-Path $subdirPath)) {
            New-Item -ItemType Directory -Path $subdirPath -Force | Out-Null
        }
    }
    
    Write-ColorOutput "✅ תיקיות נוצרו" "Green"
} catch {
    Write-ColorOutput "❌ שגיאה ביצירת תיקיות: $($_.Exception.Message)" "Red"
    exit 1
}

# התקנת תלויות Python
if (-not (Install-PythonDependencies)) {
    Write-ColorOutput "❌ התקנת תלויות נכשלה" "Red"
    exit 1
}

# העתקת קבצים
$currentPath = Get-Location
if (-not (Copy-AddinFiles -SourcePath $currentPath -DestinationPath $InstallPath)) {
    Write-ColorOutput "❌ העתקת קבצים נכשלה" "Red"
    exit 1
}

# רישום התוסף
if (-not (Register-COMAddin)) {
    Write-ColorOutput "❌ רישום התוסף נכשל" "Red"
    exit 1
}

# יצירת קיצורי דרך
if (-not (Create-Shortcuts)) {
    Write-ColorOutput "❌ יצירת קיצורי דרך נכשלה" "Red"
    exit 1
}

# יצירת סקריפט הפעלה
if (-not (Create-StartupScript)) {
    Write-ColorOutput "❌ יצירת סקריפט הפעלה נכשלה" "Red"
    exit 1
}

# בדיקת התקנה
if (-not (Test-Installation)) {
    Write-ColorOutput "❌ בדיקת התקנה נכשלה" "Red"
    exit 1
}

# סיום התקנה
Write-ColorOutput ""
Write-ColorOutput "========================================" "Green"
Write-ColorOutput "        התקנה הושלמה בהצלחה!" "Green"
Write-ColorOutput "========================================" "Green"
Write-ColorOutput ""
Write-ColorOutput "📋 מה לעשות עכשיו:" "Cyan"
Write-ColorOutput ""
Write-ColorOutput "1. 🔧 הפעל את השרת הראשי:" "Yellow"
Write-ColorOutput "   python app_with_ai.py" "Gray"
Write-ColorOutput ""
Write-ColorOutput "2. 🚀 הפעל את התוסף:" "Yellow"
Write-ColorOutput "   python outlook_com_addin.py" "Gray"
Write-ColorOutput "   או לחץ על הקיצור 'AI Email Manager'" "Gray"
Write-ColorOutput ""
Write-ColorOutput "3. 📧 פתח את Outlook ובחר מיילים לניתוח" "Yellow"
Write-ColorOutput ""
Write-ColorOutput "4. 🎯 השתמש בכפתורי ה-Ribbon החדשים" "Yellow"
Write-ColorOutput ""
Write-ColorOutput "📞 תמיכה:" "Cyan"
Write-ColorOutput "- בדוק את הלוגים ב-outlook_addin.log" "Gray"
Write-ColorOutput "- ודא שהשרת רץ על localhost:5000" "Gray"
Write-ColorOutput "- בדוק את החיבור ל-Outlook" "Gray"
Write-ColorOutput ""

if (-not $Silent) {
    Read-Host "לחץ Enter לסגירה"
}





