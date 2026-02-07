@echo off
chcp 65001 >nul
title GAFC Audit Helper - Update

echo ================================================================
echo   GAFC Audit Helper - Update
echo ================================================================
echo.

powershell -ExecutionPolicy Bypass -File "%~dp0scripts\update_audit_helper.ps1"

if errorlevel 1 (
    echo [ERROR] Update failed!
    exit /b 1
)

exit /b 0
