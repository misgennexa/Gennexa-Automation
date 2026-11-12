@echo off
title Nexgen - Setup and Run
cd /d "%~dp0"
if not exist ".venv\Scripts\streamlit.exe" (
    echo First run: setting up environment, please wait...
    call setup_nexgen.bat nopause
)
where code >nul 2>&1 && start "" code "%~dp0"
call launch_nexgen.bat
