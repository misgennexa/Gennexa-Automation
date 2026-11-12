@echo off
title Nexgen Dashboard - KILL
echo.
echo ============================================
echo   Stopping Nexgen Analytics Dashboard...
echo ============================================
echo.

set KILLED=0

for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8502 " 2^>nul') do (
    echo Killing PID %%a on port 8502...
    taskkill /PID %%a /F >nul 2>&1
    set KILLED=1
)

for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8501 " 2^>nul') do (
    echo Killing PID %%a on port 8501...
    taskkill /PID %%a /F >nul 2>&1
    set KILLED=1
)

REM Also kill any lingering streamlit.exe processes
taskkill /IM streamlit.exe /F >nul 2>&1

if %KILLED%==1 (
    echo.
    echo [OK] Nexgen server stopped successfully.
) else (
    echo.
    echo [INFO] No Nexgen server was running.
)

echo.
ping -n 2 127.0.0.1 >nul 2>&1
