@echo off
setlocal EnableExtensions
rem ============================================================
rem  IconCleaner launcher
rem  Double-click this file to run with administrator rights.
rem  (This file is intentionally ASCII-only.)
rem ============================================================
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [IconCleaner] Requesting administrator privileges...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs" >nul
    exit /b
)
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0IconCleaner.ps1"
set _rc=%errorlevel%
if not "%_rc%"=="0" (
    echo.
    echo [IconCleaner] Exited with code %_rc%
    pause
)
