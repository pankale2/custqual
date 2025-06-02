from flask import Flask, request, send_file, render_template, redirect, url_for, session, after_this_request
import pandas as pd
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
EXPORTED_PATH = os.path.join(PROCESSED_FOLDER, 'exported.xlsx')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def load_exported():
    if os.path.exists(EXPORTED_PATH):
        df = pd.read_excel(EXPORTED_PATH)
        # Filter out rows where QUESTION ID is "~{END}~" or NaN
        df = df[df["QUESTION ID"].notna()]
        df = df[df["QUESTION ID"] != "~{END}~"]
        # Only unique, valid (QUESTION ID, COUNTRY LANGUAGE) pairs
        combos = df[["QUESTION ID", "COUNTRY LANGUAGE"]].dropna().drop_duplicates()
        combo_list = combos.values.tolist()
        return df, combo_list
    return None, []

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    processed = False
    if request.method == 'POST':
        file = request.files['file']
        if file:
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)
            # Process the Excel file
            processed_path = EXPORTED_PATH
            from excel_processing import advanced_process_excel
            advanced_process_excel(filepath, processed_path)
            try:
                os.remove(filepath)
            except Exception:
                pass
            processed = True
    return render_template('index.html', processed=processed)

@app.route('/download')
def download_exported():
    if os.path.exists(EXPORTED_PATH):
        return send_file(EXPORTED_PATH, as_attachment=True, download_name="processed.xlsx")
    return "No processed file found.", 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
