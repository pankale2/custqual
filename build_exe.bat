@echo off
echo Building CustQuals Processor executable...
echo.

REM Build CustQuals Processor executable using virtual environment

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    python -m venv venv
)

REM Activate virtual environment
call venv\Scripts\activate

REM Install required packages
pip install -r requirements.txt

REM Clean previous builds
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "__pycache__" rmdir /s /q "__pycache__"

REM Build the executable
pyinstaller CustQualPro.spec

REM Check if build was successful
if exist "dist\CustQualPro.exe" (
    echo.
    echo Build successful! Executable created at: dist\CustQualPro.exe
    echo.
    echo To run the application:
    echo 1. Navigate to the dist folder
    echo 2. Run CustQualPro.exe
    echo 3. Open browser and go to http://localhost:8080
    echo.
    pause
) else (
    echo.
    echo Build failed! Please check for errors above.
    pause
)

REM Deactivate virtual environment
call venv\Scripts\deactivate
