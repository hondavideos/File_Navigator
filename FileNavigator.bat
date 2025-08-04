@echo off
REM FileNavigator.bat - Easy launcher for non-technical users
REM This bypasses PowerShell execution policy without requiring admin rights

echo Starting File Navigator...
echo.

REM Method 1: Try setting execution policy for current user (no admin required)
powershell.exe -Command "try { Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop; Write-Host 'Execution policy set successfully' -ForegroundColor Green } catch { Write-Host 'Note: Using bypass method instead' -ForegroundColor Yellow }"

REM Method 2: Launch script with execution policy bypass (works even if Method 1 fails)
cd /d "%~dp0Scripts"
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "navigator.ps1"

REM Keep window open if there's an error so user can see what happened
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo An error occurred. Press any key to close this window.
    pause >nul
)