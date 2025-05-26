from flask import Flask, request, send_file, render_template
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = '/tmp/uploads'
PROCESSED_FOLDER = '/tmp/processed'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file:
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)
            # Process the Excel file
            processed_path, processed_filename = process_excel(filepath, filename)
            response = send_file(processed_path, as_attachment=True, download_name=processed_filename)
            try:
                os.remove(filepath)
            except Exception:
                pass
            try:
                os.remove(processed_path)
            except Exception:
                pass
            return response
    return render_template('index.html')

def process_excel(filepath, filename):
    from excel_processing import advanced_process_excel
    # Insert ' - Output' before the file extension
    name, ext = os.path.splitext(filename)
    processed_filename = f"{name} - Output{ext}"
    processed_path = os.path.join(PROCESSED_FOLDER, processed_filename)
    advanced_process_excel(filepath, processed_path)
    return processed_path, processed_filename

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
