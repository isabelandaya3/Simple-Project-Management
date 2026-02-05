@echo off
REM ============================================
REM LEB Tracker - Start All
REM ============================================
REM This batch file starts the tracker server
REM and opens the dashboard in your browser.
REM ============================================

title LEB Tracker

echo.
echo ============================================
echo LEB RFI/Submittal Tracker
echo ============================================
echo.

REM Wait for network to be ready (important for startup)
REM This delay helps ensure network drives are mounted
echo Waiting for system to initialize...
timeout /t 15 /nobreak >nul

REM Navigate to script directory
cd /d "%~dp0"

REM Check if server is already running by trying to connect
powershell -Command "(New-Object Net.Sockets.TcpClient).Connect('localhost', 5000)" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Server is already running.
    echo Opening dashboard...
    start "" "http://localhost:5000"
    exit /b 0
)

REM Check if Python is available
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python not found in PATH.
    echo Please install Python and add it to your PATH.
    pause
    exit /b 1
)

REM Check if virtual environment exists
if exist "venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call venv\Scripts\activate.bat
) else if exist ".venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call .venv\Scripts\activate.bat
) else (
    echo No virtual environment found, using system Python.
)

echo Starting tracker server...
echo.

REM Start the Python application in a new minimized window
start "LEB Tracker Server" /min cmd /c "python app.py"

REM Wait a moment for server to start
echo Waiting for server to start...
timeout /t 3 /nobreak >nul

REM Open the dashboard
echo Opening dashboard...
start "" "http://localhost:5000"

echo.
echo ============================================
echo Server is running in the background.
echo Close the "LEB Tracker Server" window to stop.
echo ============================================
echo.
