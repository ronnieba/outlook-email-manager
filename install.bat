@echo off
chcp 65001 > nul
setlocal

:: =============================================================================
::  AI Email Manager - Unified Installer for Outlook COM Add-in
::  Version: 1.0
::  This script handles uninstallation of old versions, installation,
::  and verification of the Python-based COM Add-in for Outlook.
:: =============================================================================

echo.
echo  ================================================================
echo      AI Email Manager - Outlook COM Add-in Installer
echo  ================================================================
echo.
echo  This script will install the AI Email Manager add-in for Outlook.
echo  Please make sure Outlook is closed before proceeding.
echo.
pause
echo.


:: -------------------------------------------------
:: Step 1: Check prerequisites
:: -------------------------------------------------
echo [Step 1/5] Checking prerequisites...

:: Check for Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   [ERROR] Python is not installed or not in PATH.
    echo   Please install Python 3.8+ from https://www.python.org/downloads/
    goto:failure
)
echo   [OK] Python is installed.

:: Check for Outlook
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [ERROR] Microsoft Outlook is not installed.
    echo   Please install Microsoft Outlook 2016 or newer.
    goto:failure
)
echo   [OK] Microsoft Outlook is installed.
echo.


:: -------------------------------------------------
:: Step 2: Uninstall previous versions
:: -------------------------------------------------
echo [Step 2/5] Cleaning up previous versions...

:: Unregister all known python add-in files to avoid conflicts
echo   - Unregistering old COM components...
:: We run these commands even if the files don't exist, to clean up the registry.
:: Errors are ignored (>nul 2>&1) because the files might have been deleted.
python outlook_com_addin.py --unregister >nul 2>&1
python outlook_com_addin_final.py --unregister >nul 2>&1
python outlook_com_addin_simple_dll.py --unregister >nul 2>&1
python outlook_com_addin_minimal.py --unregister >nul 2>&1
python outlook_com_addin_simple_fixed.py --unregister >nul 2>&1
python outlook_com_addin_registered.py --unregister >nul 2>&1
python outlook_com_addin_working.py --unregister >nul 2>&1
python outlook_com_addin_working_final.py --unregister >nul 2>&1
python outlook_com_addin_ultra_simple.py --unregister >nul 2>&1
python outlook_simple_addin.py --unregister >nul 2>&1
python outlook_ribbon_addin.py --unregister >nul 2>&1


:: Remove the registry key from Outlook add-ins
echo   - Removing old Outlook registry entries...
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /f >nul 2>&1

echo   [OK] Cleanup complete.
echo.


:: -------------------------------------------------
:: Step 3: Install Python dependencies
:: -------------------------------------------------
echo [Step 3/5] Installing required Python packages...
pip install --upgrade pywin32 requests >nul
if %errorlevel% neq 0 (
    echo   [ERROR] Failed to install required packages.
    echo   Please install them manually: pip install pywin32 requests
    goto:failure
)
echo   [OK] pywin32 and requests are up to date.
echo.


:: -------------------------------------------------
:: Step 4: Register the new COM Add-in
:: -------------------------------------------------
echo [Step 4/5] Registering the new add-in...

:: We will use 'outlook_com_addin.py' as the definitive add-in file.
set ADDIN_FILE=outlook_com_addin.py

if not exist "%ADDIN_FILE%" (
    echo   [ERROR] Add-in file not found: %ADDIN_FILE%
    echo   Please ensure the file exists in the current directory.
    goto:failure
)

echo   - Registering %ADDIN_FILE% as a COM server...
python %ADDIN_FILE% --register
if %errorlevel% neq 0 (
    echo   [ERROR] Failed to register the COM add-in.
    echo   Try running this script as an Administrator.
    goto:failure
)
echo   [OK] COM component registered successfully.

echo   - Adding add-in to Outlook registry...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "FriendlyName" /t REG_SZ /d "AI Email Manager" /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "Description" /t REG_SZ /d "AI-powered email and meeting analysis for Outlook" /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\AIEmailManager.Addin" /v "CommandLineSafe" /t REG_DWORD /d 0 /f >nul
echo   [OK] Add-in configured in Outlook.
echo.


:: -------------------------------------------------
:: Step 5: Finalizing
:: -------------------------------------------------
echo [Step 5/5] Installation complete!
echo.
echo  ================================================================
echo                      SUCCESS!
echo  ================================================================
echo.
echo  What to do next:
echo.
echo  1. Start the main server by running:
echo     ^> python app_with_ai.py
echo.
echo  2. Open Outlook. The "AI Email Manager" add-in should now be active.
echo     To verify, go to: File ^> Options ^> Add-ins.
echo     "AI Email Manager" should appear in the "Active Application Add-ins" list.
echo.
goto:end

:failure
echo.
echo  ================================================================
echo                      INSTALLATION FAILED
echo  ================================================================
echo.
echo  Please review the error messages above and try again.

:end
echo Press any key to close...
pause > nul
endlocal