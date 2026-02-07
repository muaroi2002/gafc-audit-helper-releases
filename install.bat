@echo off
chcp 65001 >nul
title GAFC Audit Helper - Installation

echo ================================================================
echo   GAFC Audit Helper - Installation
echo ================================================================
echo.

REM Run PowerShell script
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\install_audit_helper.ps1"

if errorlevel 1 (
    echo.
    echo [ERROR] Installation failed!
    echo.
    pause
    exit /b 1
)

echo.
echo Installation completed! You can close this window.
echo.
pause
exit /b 0
