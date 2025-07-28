# GitHub Copilot Instructions for CustQuals Processor

- After th eprompt is provided, do not directly make changes to the code or project files.
- Provide the plan first. Ask any queries or confusions or difficult choices.
- Always ask before removing any code that can remove any existing functionality. 


## Project Overview
This is a Flask web application that processes Excel files from SSRS containing custom qualification data. The app:
- Processes Excel files in memory without requiring file system access
- Applies formatting, calculations, and language detection to Excel files
- Uses a modern, clean UI with a blue gradient theme
- Can be packaged as a standalone executable with PyInstaller

## Code Style Guidelines
- Use clear, descriptive variable and function names
- Add comments for complex logic or business rules
- Format Excel files consistently (column widths, data types, etc.)
- Optimize for performance, especially with large Excel files
- Prefer in-memory processing over file system operations
- Use Python type hints where possible

## Architecture
- **app.py**: Main Flask application with routes and web server
- **excel_processing.py**: Excel processing functions (pandas/openpyxl)
- **templates/**: HTML templates using Jinja2
- **static/**: CSS and other static assets
- **.github/**: GitHub and Copilot configuration

## Excel Processing Guidelines
- Use pandas for data manipulation and openpyxl for styling/formatting
- Prefer column-based operations over row-by-row operations
- Apply consistent conditional formatting for key data columns
- Freeze header rows and apply autofilter for better user experience
- Set appropriate column widths based on content type
- Format dates consistently (YYYY-MM-DD)

## UI Guidelines
- Use the established color scheme:
  - Primary gradient: #277bf5 → #00c6fb (blue gradient)
  - Background: #f7f8fa (light gray)
  - Text: #22223b (dark blue/gray)
- Keep the interface clean and minimal
- Ensure responsive design for different screen sizes
- Add subtle animations for better user feedback

## Testing Considerations
- Consider edge cases with different Excel file formats
- Ensure memory usage is reasonable for large files
- Test with various language text in the input files
- Validate Excel formula generation

## Deployment
- Package as executable using PyInstaller
- Ensure all dependencies are properly included
- Maintain in-memory processing to avoid file system permission issues
