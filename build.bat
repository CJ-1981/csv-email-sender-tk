@echo off
REM Build script to package CSV Email Sender as a single executable using PyInstaller
REM Usage: build.bat

echo ========================================
echo CSV Email Sender - PyInstaller Build
echo ========================================
echo.

REM Check if virtual environment exists
if not exist "venv\" (
    echo [!] Virtual environment not found. Creating one...
    python -m venv venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        echo Please make sure Python 3.8+ is installed and in your PATH.
        pause
        exit /b 1
    )
)

REM Activate virtual environment
echo [*] Activating virtual environment...
call venv\Scripts\activate.bat

REM Install requirements
echo [*] Installing dependencies...
pip install -q --upgrade pip
pip install -q -r requirements.txt

REM Install PyInstaller
echo [*] Installing PyInstaller...
pip install -q pyinstaller

REM Check if main.py exists
if not exist "main.py" (
    echo [ERROR] main.py not found in current directory.
    pause
    exit /b 1
)

REM Clean previous builds
echo [*] Cleaning previous builds...
if exist "build\" rmdir /s /q build
if exist "dist\CSV-Email-Sender" rmdir /s /q "dist\CSV-Email-Sender"

REM Build executable
echo.
echo [*] Building executable with PyInstaller...
echo This may take a minute...
echo.

pyinstaller --name="CSV-Email-Sender" ^
    --onefile ^
    --windowed ^
    --icon=NONE ^
    --add-data="config.py;." ^
    --add-data="csv_parser.py;." ^
    --add-data="email_sender.py;." ^
    --hidden-import=tkinter ^
    --hidden-import=openpyxl ^
    --hidden-import=ssl ^
    --hidden-import=smtplib ^
    --hidden-import=email ^
    --hidden-import=email.mime ^
    --hidden-import=email.mime.text ^
    --hidden-import=email.mime.multipart ^
    --hidden-import=email.mime.application ^
    --hidden-import=email.utils ^
    --hidden-import=queue ^
    --hidden-import=threading ^
    --hidden-import=random ^
    --hidden-import=time ^
    --hidden-import=os ^
    --hidden-import=csv ^
    --clean ^
    --noconfirm ^
    main.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed successfully!
echo ========================================
echo.
echo Executable location: dist\CSV-Email-Sender.exe
echo.
echo To distribute:
echo   - Copy dist\CSV-Email-Sender.exe to any Windows machine
echo   - No Python installation required
echo.

pause
