@echo off
echo Building CustQuals Processor executable...
echo.

REM Install required packages
pip install pyinstaller flask pandas openpyxl langdetect werkzeug

REM Clean previous builds
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "__pycache__" rmdir /s /q "__pycache__"

REM Build the executable
pyinstaller --clean CustQualPro.spec --add-data "static/css/main.css;static/css" --add-data "static/js/app.js;static/js" --add-data "static/favicon.ico;static" --add-data "templates;templates" --hidden-import "openpyxl" --hidden-import "flask" --hidden-import "werkzeug" run.py

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
