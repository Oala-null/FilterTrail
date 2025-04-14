@echo off
echo ================================================
echo FilterTrail Override Build Process
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

echo Creating hook overrides for pydantic...
if not exist "hooks" mkdir hooks
echo # Empty hook for pydantic to override the default > hooks\hook-pydantic.py
echo # This prevents PyInstaller from trying to analyze pydantic >> hooks\hook-pydantic.py
echo. >> hooks\hook-pydantic.py
echo # Define empty lists to prevent issues >> hooks\hook-pydantic.py
echo hiddenimports = [] >> hooks\hook-pydantic.py
echo excludedimports = [] >> hooks\hook-pydantic.py
echo   [OK] Created pydantic hook override

echo Building FilterTrail override version...
python -m PyInstaller ^
    --name=FilterTrail_Override ^
    --windowed ^
    --onedir ^
    --clean ^
    --add-data="filter_data.json;." ^
    --additional-hooks-dir=hooks ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=PyQt5.QtWebEngineWidgets ^
    --exclude-module=pydantic ^
    --exclude-module=nltk ^
    --noupx ^
    main.py

if exist "dist\FilterTrail_Override" (
    echo.
    echo [SUCCESS] Override build completed successfully!
    echo Application folder created at: %CD%\dist\FilterTrail_Override
    echo.
    echo This version uses hook overrides to solve Anaconda environment issues.
    echo.
    echo To run the application:
    echo 1. Navigate to: %CD%\dist\FilterTrail_Override
    echo 2. Run the FilterTrail_Override.exe executable
    echo 3. Make sure Excel is open with a worksheet
) else (
    echo.
    echo [ERROR] Build failed! Application folder not found.
    echo Please check the error messages above.
)

echo.
echo ================================================
pause