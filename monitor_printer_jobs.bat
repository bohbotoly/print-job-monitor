
@echo off
title monitor_printer_jobs_V4.3 
echo ************************************
echo monitor_printer_jobs_V4.3
echo            by bohbotoly
echo ************************************
:START
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0monitor_printer_jobs.ps1"
goto start