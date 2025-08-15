from flask import Flask, request, send_file, render_template
import os
import time
from werkzeug.utils import secure_filename
import sys
import io
from excel_processing import advanced_process_excel_memory
import webbrowser
import threading

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
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
            
            # Immediately return the processed file for download
            return send_file(
                io.BytesIO(processed_file_data),
                as_attachment=True,
                download_name=processed_filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    
    return render_template('index.html')

def open_browser():
    """Open browser after a short delay to ensure server is running"""
    time.sleep(1.5)
    webbrowser.open('http://localhost:8080')

if __name__ == '__main__':
    print("CustQuals Processor starting...")
    print("Opening browser automatically...")
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()
    print("Server running at: http://localhost:8080")
    print("Press Ctrl+C to stop the server")
    try:
        app.run(host='0.0.0.0', port=8080, debug=False)
    except KeyboardInterrupt:
        print("\nServer stopped by user (Ctrl+C).")
        sys.exit(0)