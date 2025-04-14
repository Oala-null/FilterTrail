@echo off
echo ================================================
echo FilterTrail Build Process
echo ================================================

echo Creating empty filter data file...
echo [] > filter_data.json
echo   [OK] Created empty filter_data.json

echo Cleaning build artifacts...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del *.spec
if exist "__pycache__" rmdir /s /q __pycache__
echo   [OK] Removed old build files
echo Cleanup complete.

REM Check if resources directory exists
if not exist "resources" mkdir resources
echo   [OK] Checked resources directory

REM Check for icon file
set "ICON_PARAM="
if exist "resources\favicon.ico" (
    set "ICON_PARAM=--icon=resources\favicon.ico"
    echo   [OK] Found application icon
)

echo Building FilterTrail executable...
python -m PyInstaller ^
    --name=FilterTrail ^
    --windowed ^
    --onefile ^
    --clean ^
    --add-data="filter_data.json;." ^
    --hidden-import=win32com ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=PyQt5.QtWebEngineWidgets ^
    --exclude-module=matplotlib ^
    --exclude-module=scipy ^
    --exclude-module=pandas.plotting ^
    --exclude-module=holoviews ^
    --exclude-module=numpy.random ^
    --exclude-module=pydantic ^
    --exclude-module=IPython ^
    --exclude-module=PySide2 ^
    --exclude-module=nltk ^
    --upx-dir=. ^
    main.py

if exist "dist\FilterTrail.exe" (
    echo.
    echo [SUCCESS] Build completed successfully!
    echo Executable created at: %CD%\dist\FilterTrail.exe
    echo Note: This executable uses CDN for Plotly to reduce file size.
    echo.
    echo To run the application:
    echo 1. Make sure Excel is open with a worksheet
    echo 2. Double-click FilterTrail.exe
    echo 3. Click 'Start Monitoring' and apply filters in Excel
) else (
    echo.
    echo [ERROR] Build failed! Executable not found.
    echo Please check the error messages above.
)

echo.
echo ================================================
pause