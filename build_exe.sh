# PyInstaller build script for CustQuals Processor
# Run this command from the project root directory

pyinstaller --onefile \
    --windowed \
    --name "CustQuals_Processor" \
    --icon "static/favicon.ico" \
    --add-data "static/css/main.css:static/css" \
    --add-data "static/js/app.js:static/js" \
    --add-data "static/favicon.ico:static" \
    --add-data "templates:templates" \
    --hidden-import "openpyxl" \
    --hidden-import "flask" \
    --hidden-import "werkzeug" \
    run.py

# Ensure all necessary files are included for Linux/Mac builds.