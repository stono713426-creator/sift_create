@echo off
cd /d "%~dp0"

echo ============================================================
echo   Shift Generator - Initial Setup
echo ============================================================
echo.

where py >nul 2>nul
if %errorlevel% equ 0 (
    echo [INFO] Found Python launcher (py). Creating virtual environment with Python 3.12...
    py -3.12 -m venv .venv 2>nul
    if not exist .venv\Scripts\python.exe (
        echo [INFO] Python 3.12 not found via launcher. Trying default python...
        py -m venv .venv
    )
) else (
    where python >nul 2>nul
    if %errorlevel% equ 0 (
        echo [INFO] Creating virtual environment with system python...
        python -m venv .venv
    ) else (
        echo [ERROR] Python is not installed on this PC.
        echo Please install Python 3.12 from https://www.python.org/
        echo (Check 'Add Python to PATH' during installation!)
        pause
        exit /b 1
    )
)

if not exist .venv\Scripts\python.exe (
    echo [ERROR] Failed to create virtual environment .venv.
    pause
    exit /b 1
)

echo.
echo [INFO] Installing required packages from requirements.txt...
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo [ERROR] Package installation failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Setup completed successfully!
echo   You can now launch the app via 'start_gui.bat'.
echo ============================================================
pause
