@echo off
SETLOCAL EnableDelayedExpansion

echo ==================================================
echo   PDF Table Extractor - Launcher
echo ==================================================
echo.

:: 1. Check if Python is installed (check python and python3)
set PYTHON_CMD=
for %%P in (python python3) do (
    where %%P >nul 2>&1
    if !errorlevel! == 0 (
        set PYTHON_CMD=%%P
        goto :python_found
    )
)

:python_check_failed
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

:python_found
:: 2. Check and Install Dependencies if missing
echo [*] Checking dependencies...
!PYTHON_CMD! -c "import google.genai" >nul 2>&1
if !errorlevel! neq 0 (
    echo [*] Installing dependencies...
    !PYTHON_CMD! -m pip install --upgrade pip >nul 2>&1
    !PYTHON_CMD! -m pip install -r src/requirements.txt >nul 2>&1
    if !errorlevel! neq 0 (
        echo [X] Failed to install dependencies.
        pause
        exit /b 1
    )
    echo [OK] Dependencies installed.
) else (
    echo [OK] Dependencies are up to date.
)

:: 3. Create Desktop Shortcut (if not exists)
set "SC_NAME=PDF Table Extractor.lnk"
for /f "usebackq tokens=*" %%D in (`powershell -Command "[Environment]::GetFolderPath('Desktop')"`) do set "SC_PATH=%%D\%SC_NAME%"

    echo [*] Ensuring Icon exists...
    !PYTHON_CMD! "%~dp0src\create_ico.py" >nul 2>&1
    
    echo [*] Updating Desktop Shortcut...
    :: Get full path to pythonw based on our found python command
    for /f "delims=" %%I in ('where !PYTHON_CMD!') do (
        set "PYTHON_EXE=%%I"
        set "PYTHONW_EXE=!PYTHON_EXE:python.exe=pythonw.exe!"
        set "PYTHONW_EXE=!PYTHONW_EXE:python3.exe=pythonw3.exe!"
        :: Verify pythonw exists, otherwise fallback to python
        if not exist "!PYTHONW_EXE!" set "PYTHONW_EXE=!PYTHON_EXE!"
        goto :found_exe
    )
    :found_exe
    
    powershell -ExecutionPolicy Bypass -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%SC_PATH%'); $s.TargetPath='!PYTHONW_EXE!'; $s.Arguments='-m src.main'; $s.WorkingDirectory='%~dp0'; $s.IconLocation='%~dp0src\assets\icons\pdf_to_excel.ico'; $s.Description='Extract tables from PDF using AI'; $s.Save()"
    echo [OK] Shortcut updated.

:: 4. Launch GUI
echo [*] Launching Application...
:: Find pythonw again for launching
for /f "delims=" %%I in ('where !PYTHON_CMD!') do (
    set "EXE=%%I"
    set "PW=!EXE:python.exe=pythonw.exe!"
    set "PW=!PW:python3.exe=pythonw3.exe!"
    if exist "!PW!" (
        start "" "!PW!" -m src.main
    ) else (
        start "" "!EXE!" -m src.main
    )
    goto :launched
)
:launched

timeout /t 3 >nul
exit /b 0
