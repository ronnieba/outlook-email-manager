@echo off
chcp 65001 >nul
echo.
echo ========================================
echo    Compiling COM Add-in DLL
echo ========================================
echo.

echo Step 1: Checking for Visual Studio...
echo.

:: Check for Visual Studio C# compiler
where csc >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ C# compiler not found
    echo Please install Visual Studio or .NET Framework SDK
    echo.
    echo Alternative: Download .NET Framework SDK from Microsoft
    echo https://dotnet.microsoft.com/download/dotnet-framework
    pause
    exit /b 1
)
echo ✅ C# compiler found

echo.
echo Step 2: Compiling COM add-in...
echo.

:: Compile the C# code to DLL
csc /target:library /out:AIEmailManager.dll /reference:"C:\Program Files\Microsoft Office\Root\Office16\Microsoft.Office.Interop.Outlook.dll" outlook_addin_cs.cs
if %errorLevel% neq 0 (
    echo ❌ Compilation failed
    echo.
    echo Trying without Office reference...
    csc /target:library /out:AIEmailManager.dll outlook_addin_cs.cs
    if %errorLevel% neq 0 (
        echo ❌ Compilation failed completely
        pause
        exit /b 1
    )
)
echo ✅ COM add-in compiled successfully

echo.
echo Step 3: Registering COM add-in...
echo.

:: Register the DLL
regasm AIEmailManager.dll /codebase
if %errorLevel% neq 0 (
    echo ❌ Failed to register COM add-in
    pause
    exit /b 1
)
echo ✅ COM add-in registered successfully

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

:: Check if DLL was created
if exist AIEmailManager.dll (
    echo ✅ DLL file created: AIEmailManager.dll
) else (
    echo ❌ DLL file not found
)

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
echo    - Select "COM Add-ins" and click "Go..."
echo    - Check that "AI Email Manager" appears
echo    - Ensure it's checked (enabled)
echo.
echo 2. 🎯 Use the add-in:
echo    - Add-in will run automatically when you open Outlook
echo    - Check logs in outlook_addin_success.log
echo    - If errors occur, check outlook_addin_error.log
echo.
echo 📞 Support:
echo - DLL file: AIEmailManager.dll
echo - Check logs in outlook_addin_success.log
echo - If errors occur, check outlook_addin_error.log
echo - Ensure Outlook is running
echo.
echo Press any key to close...
pause >nul




