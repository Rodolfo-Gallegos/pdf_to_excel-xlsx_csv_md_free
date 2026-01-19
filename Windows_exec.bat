@echo off
SETLOCAL EnableDelayedExpansion

echo ==================================================
echo   PDF Table Extractor - Launcher
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
)

:: 2. Install requirements from src/
echo [*] Checking dependencies...
python -m pip install --upgrade pip >nul 2>&1
pip install -r src/requirements.txt >nul 2>&1
if !errorlevel! neq 0 (
    echo [X] Failed to install dependencies.
    pause
    exit /b 1
)

:: 3. Create Desktop Shortcut (if not exists)
set "SC_NAME=PDF Table Extractor.lnk"
set "SC_PATH=%USERPROFILE%\Desktop\%SC_NAME%"

if not exist "%SC_PATH%" (
    echo [*] Creating Desktop Shortcut...
    powershell -ExecutionPolicy Bypass -Command ^
        "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%SC_PATH%'); ^
        $s.TargetPath='pythonw.exe'; ^
        $s.Arguments='\"%~dp0src\main.py\"'; ^
        $s.WorkingDirectory='%~dp0'; ^
        $s.IconLocation='%~dp0src\assets\icons\pdf_to_excel.png'; ^
        $s.Description='Extract tables from PDF using AI'; ^
        $s.Save()"
    echo [OK] Shortcut created on Desktop.
)

:: 4. Launch GUI
echo [*] Launching Application...
start pythonw src/main.py

timeout /t 3 >nul
exit /b 0
