@echo off
chcp 65001 >nul
echo.
echo ========================================
echo    Installing Real COM Add-in for Outlook
echo ========================================
echo.

echo Step 1: Checking requirements...
echo.

:: Check if .NET Framework is installed
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ .NET Framework not installed
    echo Please install .NET Framework 4.8+ from Microsoft
    pause
    exit /b 1
)
echo ✅ .NET Framework installed

:: Check Outlook
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ Microsoft Outlook not installed
    echo Please install Microsoft Outlook 2016+ before continuing
    pause
    exit /b 1
)
echo ✅ Microsoft Outlook installed

echo.
echo Step 2: Compiling COM add-in...
echo.

:: Try to compile C# add-in
where csc >nul 2>&1
if %errorLevel% equ 0 (
    echo Compiling C# COM add-in...
    csc /target:library /out:outlook_addin.dll outlook_addin.cs
    if %errorLevel% equ 0 (
        echo ✅ C# COM add-in compiled successfully
        goto :register_com
    ) else (
        echo ❌ Failed to compile C# add-in
    )
) else (
    echo ❌ C# compiler not found
)

:: Try PowerShell approach
echo Trying PowerShell approach...
powershell -Command "Add-Type -Path 'outlook_addin.cs' -OutputAssembly 'outlook_addin.dll'"
if %errorLevel% equ 0 (
    echo ✅ PowerShell compiled COM add-in successfully
    goto :register_com
) else (
    echo ❌ PowerShell compilation failed
)

:: Fallback to VBScript
echo Using VBScript fallback...
goto :register_vbs

:register_com
echo.
echo Step 3: Registering COM add-in...
echo.

:: Register the DLL
regasm outlook_addin.dll /codebase
if %errorLevel% equ 0 (
    echo ✅ COM add-in registered successfully
    goto :register_outlook
) else (
    echo ❌ Failed to register COM add-in
    goto :register_vbs
)

:register_vbs
echo.
echo Step 3: Registering VBScript COM add-in...
echo.

:: Register VBScript as COM add-in
regsvr32 /s outlook_com_addin.vbs
if %errorLevel% equ 0 (
    echo ✅ VBScript COM add-in registered successfully
) else (
    echo ❌ Failed to register VBScript COM add-in
)

:register_outlook
echo.
echo Step 4: Registering add-in in Outlook...
echo.

:: Create add-in registration in Outlook
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul 2>&1
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "FriendlyName" /t REG_SZ /d "AI Email Manager" /f >nul 2>&1
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "Description" /t REG_SZ /d "AI-powered email and meeting analysis for Outlook" /f >nul 2>&1
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "CommandLineSafe" /t REG_DWORD /d 0 /f >nul 2>&1

echo ✅ Add-in registered in Outlook

echo.
echo Step 5: Verifying installation...
echo.

:: Check add-in registration in Outlook
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ Add-in not registered in Outlook
) else (
    echo ✅ Add-in registered in Outlook
)

echo.
echo ========================================
echo           Installation Complete!
echo ========================================
echo.
echo 📋 What to do next:
echo.
echo 1. 📧 Open Outlook and check:
echo    - File → Options → Add-ins
echo    - Check that "AI Email Manager" appears
echo    - Ensure it's checked (enabled)
echo    - Check for no runtime errors
echo.
echo 2. 🎯 Use the add-in:
echo    - Add-in will run automatically when you open Outlook
echo    - Check logs in outlook_addin_success.log
echo    - If errors occur, check outlook_addin_error.log
echo.
echo 📞 Support:
echo - Check logs in outlook_addin_success.log
echo - If errors occur, check outlook_addin_error.log
echo - Ensure Outlook is running
echo - Check COM registration
echo.
echo Press any key to close...
pause >nul




