@echo off
SETLOCAL EnableDelayedExpansion

echo ==================================================
echo   PDF Table Extractor - Windows Setup
echo ==================================================
echo.

:: 1. Check if Python is installed
python --version >nul 2>&1
if !errorlevel! neq 0 (
    echo [!] Python not found. Attempting to install via winget...
    winget install --id Python.Python.3.11 --exact --silent --accept-package-agreements --accept-source-agreements
    if !errorlevel! neq 0 (
        echo [X] Automated installation failed.
        echo Please install Python manually from https://www.python.org/downloads/
        pause
        exit /b 1
    )
    echo [OK] Python installed successfully. Please restart this script.
    pause
    exit /b 0
) else (
    echo [OK] Python is already installed.
)

:: 2. Upgrade pip
echo.
echo [*] Upgrading pip...
python -m pip install --upgrade pip

:: 3. Install requirements
echo.
echo [*] Installing dependencies from requirements.txt...
pip install -r requirements.txt
if !errorlevel! neq 0 (
    echo [X] Failed to install dependencies.
    pause
    exit /b 1
)
echo [OK] Dependencies installed.

:: 4. Launch GUI
echo.
echo [*] Launching PDF Table Extractor GUI...
start pythonw gui_app.py

echo.
echo ==================================================
echo   Setup Complete! The application is starting.
echo ==================================================
echo.
timeout /t 5
exit /b 0
