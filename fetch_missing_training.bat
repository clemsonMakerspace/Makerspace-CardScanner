@echo off
REM ============================================================
REM  Makerspace Card Scanner — Fetch Missing Training Data
REM  Finds every user in the database / Excel sheet that has no
REM  Bridge LMS training record, fetches it from the API, then
REM  syncs the results back to both the database and Excel.
REM
REM  Usage:
REM    fetch_missing_training.bat            (normal run)
REM    fetch_missing_training.bat --dry-run  (preview only, no API calls)
REM    fetch_missing_training.bat --no-sync  (skip final sync step)
REM
REM  All arguments are forwarded to fetch_missing_training.py
REM ============================================================

setlocal

REM Move to the directory containing this batch file so relative
REM paths (hardware_users.db, hardware_users.xlsx, config.py …)
REM resolve correctly regardless of where the script is launched from.
cd /d "%~dp0"

echo.
echo ============================================================
echo  Makerspace -- Fetch Missing Training Data
echo ============================================================
echo.

REM Check that Python is on the PATH
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found on PATH.
    echo Make sure Python 3.8+ is installed and added to your PATH.
    echo.
    pause
    exit /b 1
)

REM Check that the driver script exists
if not exist "fetch_missing_training.py" (
    echo ERROR: fetch_missing_training.py not found.
    echo Make sure you are running this from the project root directory.
    echo.
    pause
    exit /b 1
)

REM Check that config.py exists (required for API credentials)
if not exist "config.py" (
    echo WARNING: config.py not found.
    echo The Bridge API will not be called without BRIDGE_API_URL and
    echo BRIDGE_AUTH_TOKEN set in config.py.
    echo.
)

REM Run the Python script, forwarding any extra arguments
REM e.g.  fetch_missing_training.bat --dry-run
python fetch_missing_training.py %*

REM ---- Result handling ----
if errorlevel 1 (
    echo.
    echo ============================================================
    echo  Script finished with errors (exit code %errorlevel%).
    echo  Check the output above for details.
    echo ============================================================
    echo.
    pause
    exit /b %errorlevel%
) else (
    echo.
    echo ============================================================
    echo  All done! Training data updated and synced.
    echo ============================================================
    echo.
    REM Auto-close after 5 seconds when run without errors so it
    REM doesn't block automated schedules.  Remove these two lines
    REM if you prefer the window to stay open.
    timeout /t 5 /nobreak >nul
)

endlocal
