@echo off
REM CSV Email Sender - Windows Launch Script

cd /d "%~dp0"

REM Check if venv exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    echo Installing dependencies...
    venv\Scripts\pip install -r requirements.txt
)

REM Check if tkinter is available
venv\Scripts\python -c "import tkinter" 2>nul
if errorlevel 1 (
    echo ERROR: Tkinter is not available in this Python installation.
    echo.
    echo Tkinter should be included with Python. Try reinstalling Python from:
    echo https://www.python.org/downloads/
    echo.
    echo Make sure to check "tcl/tk and IDLE" during installation.
    pause
    exit /b 1
)

echo Starting CSV Email Sender...
venv\Scripts\python main.py
pause
