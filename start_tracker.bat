@echo off
REM ============================================
REM LEB Tracker - Start Tracker Server
REM ============================================
REM This batch file starts the Python backend server.
REM The server will run on http://localhost:5000
REM ============================================

title LEB Tracker Server

echo.
echo ============================================
echo LEB RFI/Submittal Tracker
echo ============================================
echo.

REM Check if Python is available
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python not found in PATH.
    echo Please install Python and add it to your PATH.
    pause
    exit /b 1
)

REM Navigate to script directory
cd /d "%~dp0"

REM Check if virtual environment exists
if exist "venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call venv\Scripts\activate.bat
) else if exist ".venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call .venv\Scripts\activate.bat
) else (
    echo No virtual environment found, using system Python.
    echo Consider creating one with: python -m venv venv
)

echo.
echo Starting tracker server...
echo.
echo Dashboard will be available at: http://localhost:5000
echo Press Ctrl+C to stop the server.
echo.

REM Start the Python application
python app.py

REM If we get here, the server has stopped
echo.
echo Server stopped.
pause
