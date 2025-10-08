# AI Email Manager - הסרת תוסף COM
# PowerShell Script for COM Add-in Uninstallation

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

function Remove-COMAddin {
    Write-ColorOutput "🔧 מסיר תוסף COM מ-Windows Registry..." "Yellow"
    
    try {
        # הסרת רישום התוסף
        $regPaths = @(
            "HKCU:\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin",
            "HKCU:\Software\AIEmailManager",
            "HKCU:\Software\Classes\AIEmailManager.Addin",
            "HKCU:\Software\Classes\CLSID\{12345678-1234-1234-1234-123456789012}"
        )
        
        foreach ($regPath in $regPaths) {
            if (Test-Path $regPath) {
                Remove-Item -Path $regPath -Recurse -Force
                Write-ColorOutput "  ✅ הוסר: $regPath" "Green"
            }
        }
        
        Write-ColorOutput "✅ תוסף הוסר מ-Windows Registry" "Green"
        return $true
        
    } catch {
        Write-ColorOutput "❌ שגיאה בהסרת התוסף: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Remove-Shortcuts {
    Write-ColorOutput "🔗 מסיר קיצורי דרך..." "Yellow"
    
    try {
        $desktop = [Environment]::GetFolderPath("Desktop")
        $startMenu = [Environment]::GetFolderPath("StartMenu")
        
        # הסרת קיצור דרך משולחן העבודה
        $desktopShortcut = Join-Path $desktop "AI Email Manager.lnk"
        if (Test-Path $desktopShortcut) {
            Remove-Item $desktopShortcut -Force
            Write-ColorOutput "  ✅ הוסר קיצור דרך משולחן העבודה" "Green"
        }
        
        # הסרת קיצור דרך מתפריט התחל
        $startShortcut = Join-Path $startMenu "AI Email Manager.lnk"
        if (Test-Path $startShortcut) {
            Remove-Item $startShortcut -Force
            Write-ColorOutput "  ✅ הוסר קיצור דרך מתפריט התחל" "Green"
        }
        
        Write-ColorOutput "✅ קיצורי דרך הוסרו" "Green"
        return $true
        
    } catch {
        Write-ColorOutput "❌ שגיאה בהסרת קיצורי דרך: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Remove-InstallationFiles {
    param([string]$InstallPath)
    
    Write-ColorOutput "📁 מסיר קבצי התקנה..." "Yellow"
    
    try {
        if (Test-Path $InstallPath) {
            if ($Force) {
                Remove-Item -Path $InstallPath -Recurse -Force
                Write-ColorOutput "✅ כל קבצי ההתקנה הוסרו" "Green"
            } else {
                # הסרה סלקטיבית של קבצי התוסף בלבד
                $addinFiles = @(
                    "outlook_com_addin.py",
                    "outlook_addin\manifest.xml",
                    "outlook_addin\taskpane.html",
                    "outlook_addin\taskpane.js", 
                    "outlook_addin\taskpane.css",
                    "start_addin.bat",
                    "outlook_addin.log"
                )
                
                foreach ($file in $addinFiles) {
                    $filePath = Join-Path $InstallPath $file
                    if (Test-Path $filePath) {
                        Remove-Item $filePath -Force
                        Write-ColorOutput "  ✅ הוסר: $file" "Green"
                    }
                }
                
                Write-ColorOutput "✅ קבצי התוסף הוסרו" "Green"
            }
        } else {
            Write-ColorOutput "⚠️ תיקיית התקנה לא נמצאה" "Yellow"
        }
        
        return $true
        
    } catch {
        Write-ColorOutput "❌ שגיאה בהסרת קבצים: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Cleanup-OutlookCache {
    Write-ColorOutput "🧹 מנקה מטמון Outlook..." "Yellow"
    
    try {
        # ניסיון לסגור את Outlook אם פתוח
        $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        if ($outlookProcesses) {
            Write-ColorOutput "  ⚠️ Outlook פתוח, אנא סגור אותו ידנית" "Yellow"
            Write-ColorOutput "  ואז הפעל מחדש את הסקריפט" "Yellow"
            return $false
        }
        
        # ניקוי מטמון Outlook
        $cachePaths = @(
            "$env:LOCALAPPDATA\Microsoft\Outlook",
            "$env:APPDATA\Microsoft\Outlook"
        )
        
        foreach ($cachePath in $cachePaths) {
            if (Test-Path $cachePath) {
                # ניקוי קבצי מטמון של תוספים
                $cacheFiles = Get-ChildItem -Path $cachePath -Filter "*AIEmailManager*" -ErrorAction SilentlyContinue
                foreach ($file in $cacheFiles) {
                    Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
                }
            }
        }
        
        Write-ColorOutput "✅ מטמון Outlook נוקה" "Green"
        return $true
        
    } catch {
        Write-ColorOutput "❌ שגיאה בניקוי מטמון: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Show-UninstallSummary {
    Write-ColorOutput ""
    Write-ColorOutput "========================================" "Green"
    Write-ColorOutput "        הסרה הושלמה בהצלחה!" "Green"
    Write-ColorOutput "========================================" "Green"
    Write-ColorOutput ""
    Write-ColorOutput "📋 מה הוסר:" "Cyan"
    Write-ColorOutput ""
    Write-ColorOutput "✅ תוסף COM מ-Windows Registry" "Green"
    Write-ColorOutput "✅ קיצורי דרך משולחן העבודה ותפריט התחל" "Green"
    Write-ColorOutput "✅ קבצי התוסף" "Green"
    Write-ColorOutput "✅ מטמון Outlook" "Green"
    Write-ColorOutput ""
    Write-ColorOutput "📝 הערות חשובות:" "Cyan"
    Write-ColorOutput ""
    Write-ColorOutput "• הפעל מחדש את Outlook כדי להשלים ההסרה" "Yellow"
    Write-ColorOutput "• אם תרצה להתקין שוב, הרץ install_com_addin.ps1" "Yellow"
    Write-ColorOutput "• קבצי הנתונים (email_manager.db) נשמרו" "Yellow"
    Write-ColorOutput ""
}

# התחלת הסרה
Write-ColorOutput ""
Write-ColorOutput "========================================" "Red"
Write-ColorOutput "   AI Email Manager - הסרת תוסף COM" "Red"
Write-ColorOutput "========================================" "Red"
Write-ColorOutput ""

# בדיקת הרשאות מנהל
if (-not (Test-Administrator)) {
    Write-ColorOutput "❌ נדרשות הרשאות מנהל להסרה" "Red"
    Write-ColorOutput "הפעל את הסקריפט כמנהל (Run as Administrator)" "Yellow"
    if (-not $Silent) {
        Read-Host "לחץ Enter לסגירה"
    }
    exit 1
}

Write-ColorOutput "✅ הרשאות מנהל מאושרות" "Green"

# אישור הסרה
if (-not $Force -and -not $Silent) {
    Write-ColorOutput ""
    Write-ColorOutput "⚠️ זה יסיר את תוסף AI Email Manager מ-Outlook" "Yellow"
    Write-ColorOutput "האם אתה בטוח שברצונך להמשיך?" "Yellow"
    $confirmation = Read-Host "הקלד 'yes' לאישור"
    if ($confirmation -ne "yes") {
        Write-ColorOutput "הסרה בוטלה" "Yellow"
        exit 0
    }
}

Write-ColorOutput ""
Write-ColorOutput "🚀 מתחיל הסרת התוסף..." "Yellow"

# הסרת רישום COM
if (-not (Remove-COMAddin)) {
    Write-ColorOutput "❌ הסרת רישום COM נכשלה" "Red"
    exit 1
}

# הסרת קיצורי דרך
if (-not (Remove-Shortcuts)) {
    Write-ColorOutput "❌ הסרת קיצורי דרך נכשלה" "Red"
    exit 1
}

# הסרת קבצי התקנה
if (-not (Remove-InstallationFiles -InstallPath $InstallPath)) {
    Write-ColorOutput "❌ הסרת קבצי התקנה נכשלה" "Red"
    exit 1
}

# ניקוי מטמון Outlook
if (-not (Cleanup-OutlookCache)) {
    Write-ColorOutput "⚠️ ניקוי מטמון Outlook נכשל" "Yellow"
    Write-ColorOutput "אנא סגור את Outlook והפעל מחדש" "Yellow"
}

# הצגת סיכום
Show-UninstallSummary

if (-not $Silent) {
    Read-Host "לחץ Enter לסגירה"
}





