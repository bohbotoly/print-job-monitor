setlocal enabledelayedexpansion
@echo off

:: Print Job Monitoring System
:: Created by bohbotoly
:: Version 4.3
:: https://github.com/bohbotoly/print-job-monitor

title Print Job Monitoring System v4.3

:: Set colors for console output
color 0B

:: Get the current date and time for logging
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YYYY=%dt:~0,4%"
set "MM=%dt:~4,2%"
set "DD=%dt:~6,2%"
set "HH=%dt:~8,2%"
set "Min=%dt:~10,2%"
set "Sec=%dt:~12,2%"
set "timestamp=%YYYY%-%MM%-%DD% %HH%:%Min%:%Sec%"

:: Display header
echo ╔════════════════════════════════════════════════════════╗
echo ║                                                        ║
echo ║           Print Job Monitoring System v4.3             ║
echo ║                   by bohbotoly                         ║
echo ║                                                        ║
echo ╚════════════════════════════════════════════════════════╝
echo.
echo Started at: %timestamp%
echo.
echo [INFO] Initializing monitoring service...

:: Check if running as administrator
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARNING] This script is not running with administrative privileges.
    echo [WARNING] Some features may not work correctly.
    echo [WARNING] Right-click the BAT file and select "Run as administrator" for full functionality.
    echo.
    timeout /t 5 >nul
)

:START
echo [INFO] Starting monitoring script at %time%
echo [INFO] Press CTRL+C to stop the monitoring service

:: Run the PowerShell script with appropriate parameters
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0print-job-monitor.ps1"

:: If the script exits, restart it automatically
echo.
echo [WARNING] Script terminated unexpectedly.
echo [INFO] Restarting in 5 seconds...
timeout /t 5 >nul
goto START

:: This section won't be reached due to the goto loop above, but included for completeness
:END
echo [INFO] Monitoring service stopped.
echo [INFO] Press any key to exit.
pause >nul
endlocal
