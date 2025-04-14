@echo off
echo Starting FilterTrail...

REM Check for available versions in priority order
if exist "dist\FilterTrail_Simple\FilterTrail_Simple.exe" (
    echo Using simple compatible version...
    start "" "dist\FilterTrail_Simple\FilterTrail_Simple.exe"
    goto :EOF
)

if exist "dist\FilterTrail_Fast\FilterTrail_Fast.exe" (
    echo Using fast directory-based version...
    start "" "dist\FilterTrail_Fast\FilterTrail_Fast.exe"
    goto :EOF
)

if exist "dist\FilterTrail.exe" (
    echo Using standard executable version...
    start "" "dist\FilterTrail.exe"
    goto :EOF
)

echo No built versions found. Using Python version...

REM Check if Python is available
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Error: Python not found. Please install Python 3.8 or newer.
    pause
    exit /b 1
)

REM Check if required modules are installed
python -c "import PyQt5" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Installing required dependencies...
    pip install -r requirements.txt
)

REM Run the application
python filter_trail.py