from flask import Flask, request, send_file, render_template, redirect, url_for, session, after_this_request
import pandas as pd
import os
import time
from werkzeug.utils import secure_filename
import sys
import io
from excel_processing import advanced_process_excel_memory
import webbrowser
import threading

app = Flask(__name__)

# Global variable to store processed file in memory
processed_file_data = None
processed_filename = None

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    global processed_file_data, processed_filename
    processed = False
    
    if request.method == 'POST':
        file = request.files['file']
        if file:
            # Generate output filename with timestamp
            timestamp = time.strftime('%Y%m%d_%H%M%S')
            base, ext = os.path.splitext(secure_filename(file.filename))
            processed_filename = f"{base}_{timestamp}.xlsx"
            
            # Read file into memory and process
            file_data = file.read()
            processed_file_data = advanced_process_excel_memory(file_data)
            processed = True
    
    return render_template('index.html', processed=processed)

@app.route('/download')
def download_exported():
    global processed_file_data, processed_filename
    
    if processed_file_data and processed_filename:
        return send_file(
            io.BytesIO(processed_file_data),
            as_attachment=True,
            download_name=processed_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    return "No processed file found.", 404

def open_browser():
    """Open browser after a short delay to ensure server is running"""
    import time
    time.sleep(1.5)  # Wait for server to start
    webbrowser.open('http://localhost:8080')

if __name__ == '__main__':
    print("CustQuals Processor starting...")
    print("Opening browser automatically...")
    
    # Start browser in a separate thread
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()
    
    print("Server running at: http://localhost:8080")
    print("Press Ctrl+C to stop the server")
    app.run(host='0.0.0.0', port=8080, debug=False)