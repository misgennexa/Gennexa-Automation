@echo off
title Nexgen Analytics Dashboard
cd /d "%~dp0"
echo.
echo ============================================
echo   Nexgen Analytics Dashboard v2.0
echo ============================================
echo.

:: Kill any existing streamlit servers first
echo Stopping any existing Nexgen server...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8502 " 2^>nul') do (
    taskkill /PID %%a /F >nul 2>&1
)
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8501 " 2^>nul') do (
    taskkill /PID %%a /F >nul 2>&1
)
taskkill /IM streamlit.exe /F >nul 2>&1

echo Starting Nexgen Dashboard...
ping -n 3 127.0.0.1 >nul 2>&1
start "" "http://localhost:8502"
echo.
echo Dashboard running at: http://localhost:8502
echo Close this window or press Ctrl+C to stop.
echo.
.\.venv\Scripts\streamlit.exe run nexgen_app.py --server.port 8502 --server.headless true --browser.gatherUsageStats false
echo.
echo [Nexgen stopped]
pause
