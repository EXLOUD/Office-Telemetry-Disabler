@echo off
setlocal enabledelayedexpansion

:: ====================================================
:: Define PowerShell paths
:: ====================================================
set "PS5_PATH=%systemdrive%\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS7_PATH=%ProgramFiles%\PowerShell\7\pwsh.exe"
set "PS7_PREVIEW_PATH=%ProgramFiles%\PowerShell\7-preview\pwsh.exe"

:: Find best PowerShell for admin elevation
set "ELEVATION_PS="
if exist "%PS7_PREVIEW_PATH%" (
    set "ELEVATION_PS=%PS7_PREVIEW_PATH%"
) else if exist "%PS7_PATH%" (
    set "ELEVATION_PS=%PS7_PATH%"
) else if exist "%PS5_PATH%" (
    set "ELEVATION_PS=%PS5_PATH%"
) else (
    echo [ERROR] No PowerShell found for elevation!
    pause
    exit /b 1
)

:: ====================================================
:: Launch with admin rights
:: ====================================================
if "%1"=="admin" goto :AdminMode

echo [INFO] Launch with admin rights...
"%ELEVATION_PS%" -Command "Start-Process cmd -ArgumentList '/c \"%~f0\" admin' -Verb RunAs"
exit /B

:AdminMode
pushd "%CD%"
CD /D "%~dp0"

:: ====================================================
:: Office Privacy and Telemetry Disabler Launcher
:: ====================================================

title Office Privacy and Telemetry Disabler Launcher

:: Set script directory
set "SCRIPT_DIR=%~dp0"

:: Initialize variables
set "PS_EXE="
set "PS_SCRIPT="
set "PS_VERSION="
set "SCRIPT_TYPE="

:: ====================================================
:: Find PowerShell Executable (пріоритет: Preview > 7 > 5)
:: ====================================================

if exist "%PS7_PREVIEW_PATH%" (
    set "PS_EXE=%PS7_PREVIEW_PATH%"
    set "PS_VERSION=PowerShell 7 Preview"
    set "PS_MAJOR=7"
    goto :found_powershell
)

if exist "%PS7_PATH%" (
    set "PS_EXE=%PS7_PATH%"
    set "PS_VERSION=PowerShell 7"
    set "PS_MAJOR=7"
    goto :found_powershell
)

if exist "%PS5_PATH%" (
    set "PS_EXE=%PS5_PATH%"
    set "PS_VERSION=PowerShell 5"
    set "PS_MAJOR=5"
    goto :found_powershell
)

echo [ERROR] No compatible PowerShell version found!
echo.
echo Please install either:
echo  - PowerShell 7 Preview (recommended)
echo  - PowerShell 7 
echo  - PowerShell 5 (Windows PowerShell)
echo.
pause
exit /b 1

:found_powershell

:: ====================================================
:: Detect Windows Version
:: ====================================================
for /f "tokens=4-5 delims=. " %%i in ('ver') do (
    set "WIN_MAJOR=%%i"
    set "WIN_MINOR=%%j"
)

:: ====================================================
:: Choose Script Based on Windows Version AND PowerShell Version
:: ====================================================
if !WIN_MAJOR! GEQ 10 (
    if !PS_MAJOR!==5 (
        set "SCRIPT_BASENAME=office_privacy_telemetry_disabler_win7+.ps1"
        set "SCRIPT_TYPE=Windows 10/11 with PowerShell 5"
    ) else (
        set "SCRIPT_BASENAME=office_privacy_telemetry_disabler.ps1"
        set "SCRIPT_TYPE=Windows 10/11 with PowerShell 7"
    )
) else (
    set "SCRIPT_BASENAME=office_privacy_telemetry_disabler_win7+.ps1"
    set "SCRIPT_TYPE=Windows 7/8/8.1"
)

:: ====================================================
:: Locate Script
:: ====================================================
set "SCRIPT_FOUND="

set "TEST_SCRIPT=%SCRIPT_DIR%!SCRIPT_BASENAME!"
if exist "!TEST_SCRIPT!" (
    set "PS_SCRIPT=!TEST_SCRIPT!"
    set "SCRIPT_FOUND=YES"
    goto :script_found
)

set "TEST_SCRIPT=%SCRIPT_DIR%script\!SCRIPT_BASENAME!"
if exist "!TEST_SCRIPT!" (
    set "PS_SCRIPT=!TEST_SCRIPT!"
    set "SCRIPT_FOUND=YES"
    set "SCRIPT_TYPE=!SCRIPT_TYPE! (from script folder)"
    goto :script_found
)

echo [ERROR] Expected script !SCRIPT_BASENAME! not found!
echo.
echo Please make sure this script exists:
echo  - !SCRIPT_BASENAME!
echo Either in the same folder as this launcher or in the 'script' subfolder.
echo.
pause
exit /b 1

:script_found

:: ====================================================
:: Display Information
:: ====================================================

echo.
echo ====================================================
echo    Office Privacy and Telemetry Disabler Launcher
echo.
echo                      by EXLOUD
echo              https://github.com/EXLOUD
echo ====================================================
echo.
echo System Information:
echo  - PowerShell: %PS_VERSION%
echo  - Script: !SCRIPT_TYPE!
echo  - Location: !PS_SCRIPT!
echo.
echo This will disable telemetry and privacy features for:
echo  - Microsoft Office 2010-2024
echo  - Office logging and telemetry
echo  - Customer Experience Improvement Program
echo  - Connected Experiences
echo  - Automatic updates and notifications
echo  - Scheduled telemetry tasks
echo.

:confirmation
set /p "CONFIRM=Do you want to continue? (Y/N): "

if /i "!CONFIRM!"=="y" goto :proceed
if /i "!CONFIRM!"=="yes" goto :proceed
if /i "!CONFIRM!"=="n" goto :cancel
if /i "!CONFIRM!"=="no" goto :cancel

echo Invalid input. Please enter Y or N.
goto :confirmation

:cancel
echo.
echo Operation cancelled by user.
pause
exit /b 0

:proceed
cls

echo.
echo [INFO] Launching Office Privacy Disabler...
echo [INFO] PowerShell: %PS_VERSION%
echo [INFO] Script: !SCRIPT_TYPE!
echo.
echo [WARNING] Administrator rights may be required for some registry changes.
echo.

cd /d "%SCRIPT_DIR%"

"%PS_EXE%" -ExecutionPolicy Bypass -NoProfile -File "!PS_SCRIPT!"

if %errorLevel% == 0 (
    echo.
    echo [SUCCESS] Script completed successfully!
    echo.
    echo Office privacy and telemetry settings have been disabled.
    echo Some changes may require restarting Office applications.
) else (
    echo.
    echo [ERROR] Script encountered errors. Exit code: %errorLevel%
    echo.
    echo This may happen if:
    echo  - Office is not installed
    echo  - Administrator rights are required
    echo  - Registry access is restricted
)

echo.
echo Press any key to exit...
pause >nul
exit /b 0
