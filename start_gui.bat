@echo off
cd /d "%~dp0"

echo [INFO] Starting Shift Generator GUI...

if not exist .venv\Scripts\python.exe (
    echo [ERROR] Virtual environment .venv not found.
    pause
    exit /b 1
)

.venv\Scripts\python.exe gui_app.py

if errorlevel 1 (
    echo.
    echo [ERROR] An error occurred during execution.
    pause
)
