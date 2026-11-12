@echo off
title Nexgen - First-time setup
cd /d "%~dp0"
echo Creating Python 3.12 virtual environment in .venv ...
where uv >nul 2>&1
if %errorlevel%==0 (
    uv venv --python 3.12 .venv
    uv pip install --python .\.venv\Scripts\python.exe -r requirements.txt
) else (
    py -3.12 -m venv .venv || (echo Python 3.12 not found. Install it from python.org or "pip install uv" first. & pause & exit /b 1)
    .\.venv\Scripts\python.exe -m pip install --upgrade pip
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
)
echo.
echo Setup complete. Run launch_nexgen.bat or press F5 in VS Code.
if "%1"=="" pause
