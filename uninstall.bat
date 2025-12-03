@echo off
chcp 65001 >nul
title GAFC Audit Helper - Uninstall

echo ================================================================
echo   GAFC Audit Helper - Uninstall
echo ================================================================
echo.

powershell -ExecutionPolicy Bypass -File "%~dp0scripts\uninstall_audit_helper.ps1"

if errorlevel 1 (
    echo [ERROR] Uninstall failed!
    exit /b 1
)

exit /b 0
