@echo off
REM ============================================================
REM  Makerspace Card Scanner - Application Launcher
REM  Runs MakerspaceSignInTablet.py using the bundled Python.
REM ============================================================

setlocal

REM Move to the directory containing this batch file
cd /d "%~dp0"

REM ---- Locate embedded Python --------------------------------
set "PYTHON_DIR=%~dp0python"
set "PYTHON_EXE=%PYTHON_DIR%\pythonw.exe"
set "PYTHON_CONSOLE=%PYTHON_DIR%\python.exe"

REM If embedded Python is missing, fall back to system Python
if not exist "%PYTHON_EXE%" (
    if not exist "%PYTHON_CONSOLE%" (
        echo Embedded Python not found in "%PYTHON_DIR%".
        echo Falling back to system Python...
        where python >nul 2>&1
        if errorlevel 1 (
            echo ERROR: Python not found. Please reinstall the application.
            pause
            exit /b 1
        )
        set "PYTHON_EXE=pythonw"
        set "PYTHON_CONSOLE=python"
    ) else (
        set "PYTHON_EXE=%PYTHON_CONSOLE%"
    )
)

REM ---- Set environment for embedded Python -------------------
set "PATH=%PYTHON_DIR%;%PYTHON_DIR%\Scripts;%PATH%"
set "PYTHONPATH=%~dp0"
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"

REM ---- Set Tcl/Tk paths for tkinter (needed by embedded Python) ---
if exist "%PYTHON_DIR%\tcl\tcl8.6" (
    set "TCL_LIBRARY=%PYTHON_DIR%\tcl\tcl8.6"
    set "TK_LIBRARY=%PYTHON_DIR%\tcl\tk8.6"
)

REM ---- Check that the main script exists ---------------------
if not exist "MakerspaceSignInTablet.py" (
    echo ERROR: MakerspaceSignInTablet.py not found!
    echo Please ensure all application files are present.
    echo.
    echo Troubleshooting:
    echo   1. Re-run the installer to repair the installation
    echo   2. Check that the file was not deleted or moved
    pause
    exit /b 1
)

REM ---- Launch the application --------------------------------
REM Use pythonw.exe for a clean experience (no console window).
REM To debug, change %PYTHON_EXE% to %PYTHON_CONSOLE% below.
start "Makerspace Card Scanner" "%PYTHON_EXE%" MakerspaceSignInTablet.py

endlocal
