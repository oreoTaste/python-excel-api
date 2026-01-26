@echo off
echo STEP 1: Google Auth Check...

python google_auth_setup.py
if %errorlevel% equ 0 (
    echo Session is valid. Starting app...
    goto RUN_APP
)

echo.
echo LOGIN REQUIRED
echo ----------------------------------------------------
echo 1. Close all Chrome windows.
echo 2. Chrome will open in debug mode.
echo ----------------------------------------------------
pause

start chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\temp\google_debug"

echo.
echo 3. Login to Google in the new Chrome window.
set /p dummy="After login, press [Enter] here..."

python google_auth_setup.py save
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to save session.
    pause
    exit /b
)

:RUN_APP
echo.
echo STEP 2: Running Hardwork App
python hardwork.py
pause