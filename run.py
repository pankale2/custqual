from flask import Flask, request, send_file, render_template
import os
import time
from werkzeug.utils import secure_filename
import sys
import io
from excel_processing import advanced_process_excel_memory
import webbrowser
import threading
import getpass

USER_SF_ID_MAP = {
    "pankaj": "005P40000029a1VIAQ",
    "shubham": "005P40000029Zv3IAE",
    "vishal": "005P400000FIKvhIAH",
    "vishesh": "005P40000029ZyHIAU",
    "shweta": "00568000005hiTfAAI",
    "nidhi": "005P40000029a7xIAA",
    "praveen": "005P40000029aBBIAY"
}

def get_sf_owner_id(username):
    uname = username.lower()
    for key, sfid in USER_SF_ID_MAP.items():
        if key in uname:
            return sfid
    return None

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
            username = getpass.getuser()
            sf_owner_id = get_sf_owner_id(username)
            processed_file_data = advanced_process_excel_memory(file_data, sf_owner_id)
            
            # Immediately return the processed file for download
            return send_file(
                io.BytesIO(processed_file_data),
                as_attachment=True,
                download_name=processed_filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    
    username = getpass.getuser()
    sf_owner_id = get_sf_owner_id(username)
    sf_owner_id_display = sf_owner_id if sf_owner_id else "SF ID Unknown"
    return render_template('index.html', username=username, sf_owner_id=sf_owner_id_display)

@app.route('/shutdown', methods=['POST'])
def shutdown():
    shutdown_server()
    return 'Server shutting down...'


def shutdown_server():
    print("Shutting down the server...")
    os._exit(0)  # Forcefully terminate the process

def open_browser():
    """Open browser after a short delay to ensure server is running"""
    time.sleep(1.5)
    webbrowser.open('http://localhost:5001')

# No changes needed for PyInstaller compatibility.
# Just ensure your main entry point is:
if __name__ == '__main__':
    print("CustQuals Processor starting...")
    print("Opening browser automatically...")
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()
    print("Server running at: http://localhost:5001")
    print("Press Ctrl+C to stop the server")
    try:
        app.run(host='0.0.0.0', port=5001, debug=False)
    except KeyboardInterrupt:
        print("\nServer stopped by user (Ctrl+C).")
        sys.exit(0)
