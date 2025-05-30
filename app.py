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

@app.route('/review', methods=['GET', 'POST'])
def review_page():
    df, combo_list = load_exported()
    if df is None or not combo_list:
        return "<h3>No data to review.</h3>"

    # Handle navigation first, completely separate from save operations
    try:
        current_idx = int(request.args.get('qidx', 0))
    except (TypeError, ValueError):
        current_idx = 0
    
    # POST: Save changes to EXPORTED_PATH and handle download/navigation
    if request.method == 'POST':
        question_id, country_language = combo_list[current_idx]
        rows = df[(df["QUESTION ID"] == question_id) & (df["COUNTRY LANGUAGE"] == country_language)]
        row_indices = [int(idx) for idx in rows.index]

        fb = request.form.get('feedback', '')
        rec = request.form.get('recommendation', '')
        if fb == "Other: ":
            other_text = request.form.get('other_feedback', '').strip()
            if len(other_text) >= 6 and len(other_text.split()) >= 2:
                fb = "Other: " + other_text

        # Save changes to EXPORTED_PATH (always the same file)
        for idx in row_indices:
            df.at[idx, "Feedback"] = fb
            df.at[idx, "Recommendation"] = rec
        df.to_excel(EXPORTED_PATH, index=False)

        # Download always serves the same file
        if request.form.get('download') == '1':
            return send_file(EXPORTED_PATH, as_attachment=True, download_name="processed.xlsx")

        # Handle navigation after save
        if request.form.get('nav_prev'):
            return redirect(url_for('review_page', qidx=max(0, current_idx - 1)))
        elif request.form.get('nav_next'):
            return redirect(url_for('review_page', qidx=min(len(combo_list) - 1, current_idx + 1)))
        return redirect(url_for('review_page', qidx=current_idx))

    # GET: Show the review page
    # Ensure index is within bounds
    current_idx = max(0, min(current_idx, len(combo_list) - 1))
    question_id, country_language = combo_list[current_idx]

    rows = df[(df["QUESTION ID"] == question_id) & (df["COUNTRY LANGUAGE"] == country_language)]
    unique_fields = [
        "SURVEY ID", "ACCOUNT NAME", "CSM OWNER NAME", "PM WHO CREATED QUAL",
        "EMAIL", "QUESTION TYPE", "QUESTION VISIBILITY", "DATE CREATED",
        # "COUNTRY LANGUAGE",  # now handled per page
        # "QUESTION TEXT LANGUAGE"  # now handled per row
    ]
    unique_data = {}
    for col in unique_fields:
        vals = rows[col].dropna().unique()
        if col in ["QUESTION ID", "SURVEY ID"]:
            vals = [str(int(float(v))) for v in vals if str(v).replace('.', '', 1).isdigit()]
        unique_data[col] = ", ".join(str(v) for v in vals)
    unique_data["QUESTION ID"] = str(int(float(question_id))) if str(question_id).replace('.', '', 1).isdigit() else str(question_id)
    unique_data["COUNTRY LANGUAGE"] = country_language

    # Only unique (ANSWER PRECODE, ANSWER OPTION TEXT) pairs for the table
    seen_pairs = set()
    answer_rows = []
    for idx, r in rows.iterrows():
        pair = (str(r.get("ANSWER PRECODE", "")), str(r.get("ANSWER OPTION TEXT", "")))
        if pair not in seen_pairs:
            seen_pairs.add(pair)
            row_data = {
                "row_index": idx,
                "QUESTION NAME": r.get("QUESTION NAME", ""),
                "QUESTION TEXT": r.get("QUESTION TEXT", ""),
                "ANSWER PRECODE": r.get("ANSWER PRECODE", ""),
                "ANSWER OPTION TEXT": r.get("ANSWER OPTION TEXT", ""),
                "Feedback": r.get("Feedback", ""),
                "Recommendation": r.get("Recommendation", ""),
                "COUNTRY LANGUAGE": r.get("COUNTRY LANGUAGE", ""),
                "QUESTION TEXT LANGUAGE": r.get("QUESTION TEXT LANGUAGE", ""),
            }
            if "QUESTION ID" in r and pd.notna(r["QUESTION ID"]) and str(r["QUESTION ID"]).replace('.', '', 1).isdigit():
                row_data["QUESTION ID"] = str(int(float(r["QUESTION ID"])))
            if "SURVEY ID" in r and pd.notna(r["SURVEY ID"]) and str(r["SURVEY ID"]).replace('.', '', 1).isdigit():
                row_data["SURVEY ID"] = str(int(float(r["SURVEY ID"])))
            answer_rows.append(row_data)
    return render_template(
        'review.html',
        qidx=current_idx,
        qid_count=len(combo_list),
        unique_data=unique_data,
        answer_rows=answer_rows,
        prev_idx=max(0, current_idx - 1),
        next_idx=min(len(combo_list) - 1, current_idx + 1),
        has_prev=current_idx > 0,
        has_next=current_idx < len(combo_list) - 1
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
