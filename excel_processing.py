# This script is for advanced Excel processing using pandas and openpyxl
# It will be imported and used in app.py for the process_excel function
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.formatting.rule import ColorScaleRule
import datetime
from langdetect import detect, LangDetectException
import io

def advanced_process_excel_memory(input_data):
    """
    Process Excel file entirely in memory without touching disk
    """
    # Read from bytes data
    df = pd.read_excel(io.BytesIO(input_data), skiprows=4)
    
    # Drop empty columns
    df = df.dropna(axis=1, how='all')
    
    # Build final column order (required, added, kept)
    preferred_order = [
        'SURVEY ID', 'ACCOUNT NAME', 'CSM OWNER NAME', 'PM WHO CREATED QUAL', 'EMAIL',
        'DATE FLAGGED', 'QUESTION TYPE', 'QUESTION VISIBILITY', 'DATE CREATED', 'COUNTRY LANGUAGE',
        'QUESTION ID', 'Feedback', 'QUESTION NAME', 'QUESTION TEXT', 'Recommendation',
        'ANSWER PRECODE', 'ANSWER OPTION TEXT'
    ]
    
    # Remove unwanted columns and ensure all preferred columns exist
    df.columns = df.columns.str.strip()
    cols_to_remove = [
        'ACCOUNT ID', 'BUSINESS UNIT ID', 'BUSINESS UNIT NAME', 'SF ACCOUNT ID',
        'REGION', 'CSM OWNER ID', 'CSM EMAIL', 'SURVEY ISSUE NAME', 'DATE/TIME CREATED'
    ]
    df = df.drop(columns=[col for col in cols_to_remove if col in df.columns], errors='ignore')
    for col in preferred_order:
        if col not in df.columns:
            df[col] = ''
    
    # Fill DATE FLAGGED with current system date for all rows
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    if 'DATE FLAGGED' in df.columns:
        df['DATE FLAGGED'] = today_str
    
    # Reindex to final order
    final_cols = [col for col in preferred_order if col in df.columns]
    df = df[final_cols]
    
    # Sort once by all keys present
    sort_keys = [col for col in ['QUESTION ID', 'SURVEY ID', 'ANSWER PRECODE'] if col in df.columns]
    if sort_keys:
        df = df.sort_values(by=sort_keys, kind='stable')
    
    # Detect language for each cell in 'QUESTION TEXT' column and add a new column 'QUESTION TEXT LANGUAGE'
    if 'QUESTION TEXT' in df.columns:
        def detect_lang_safe(text):
            try:
                return detect(str(text)) if pd.notnull(text) and str(text).strip() else ''
            except LangDetectException:
                return ''
        df['QUESTION TEXT LANGUAGE'] = df['QUESTION TEXT'].apply(detect_lang_safe)
    
    # Set Feedback/Recommendation based on language logic
    if all(col in df.columns for col in ['COUNTRY LANGUAGE', 'QUESTION TEXT LANGUAGE', 'Feedback', 'Recommendation']):
        mask_non_english = ~df['COUNTRY LANGUAGE'].astype(str).str.endswith('-ENG')
        mask_qtext_english = df['QUESTION TEXT LANGUAGE'] == 'en'
        # Where COUNTRY LANGUAGE doesn't end with -ENG and QUESTION TEXT is English
        mask_translate = mask_non_english & mask_qtext_english
        df.loc[mask_translate, 'Feedback'] = 'Language Translation Required'
        df.loc[mask_translate, 'Recommendation'] = 'Translate into correct Language'
        # Where COUNTRY LANGUAGE doesn't end with -ENG and QUESTION TEXT is not English
        mask_ok = mask_non_english & ~mask_qtext_english
        df.loc[mask_ok, 'Feedback'] = 'OK'
        df.loc[mask_ok, 'Recommendation'] = 'OK'
    
    # Save to memory buffer instead of file
    output_buffer = io.BytesIO()
    df.to_excel(output_buffer, index=False, engine='openpyxl')
    
    # Post-process with openpyxl in memory
    output_buffer.seek(0)
    wb = openpyxl.load_workbook(output_buffer)
    ws = wb.active
    
    # Disable word wrap, unmerge, set width
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=False)
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))
    
    # Set custom column widths
    col_widths = {
        'SURVEY ID': 9.57, 'ACCOUNT NAME': 9.57, 'CSM OWNER NAME': 9.57,
        'PM WHO CREATED QUAL': 9.57, 'EMAIL': 9.57, 'DATE FLAGGED': 9.57,
        'QUESTION TYPE': 9.57, 'QUESTION VISIBILITY': 4, 'DATE CREATED': 11,
        'ANSWER PRECODE': 5.71, 'QUESTION TEXT': 57.14, 'Feedback': 28.57,
        'Recommendation': 28.57, 'ANSWER OPTION TEXT': 20, 'COUNTRY LANGUAGE': 17.86,
        'QUESTION ID': 8.57
    }
    # Only set widths for columns in the first row (header)
    for cell in ws[1]:
        header = cell.value
        col_letter = get_column_letter(cell.column)
        if header in col_widths:
            ws.column_dimensions[col_letter].width = col_widths[header]
        else:
            ws.column_dimensions[col_letter].width = 20
    
    # Format DATE CREATED to show only date, hide time
    if 'DATE CREATED' in df.columns:
        date_col_idx = df.columns.get_loc('DATE CREATED') + 1
        for row in ws.iter_rows(min_row=2, min_col=date_col_idx, max_col=date_col_idx, max_row=ws.max_row):
            for cell in row:
                cell.number_format = 'yyyy-mm-dd'
    
    # Right align DATE FLAGGED & DATE CREATED
    for col_name in ['DATE FLAGGED', 'DATE CREATED']:
        if col_name in df.columns:
            col_idx = df.columns.get_loc(col_name) + 1
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx, max_row=ws.max_row):
                for cell in row:
                    cell.alignment = Alignment(horizontal='right')
    
    # Conditional formatting for SURVEY ID, QUESTION ID, DATE CREATED
    for col_name in ['SURVEY ID', 'QUESTION ID', 'DATE CREATED']:
        if col_name in df.columns:
            col_idx = df.columns.get_loc(col_name) + 1
            col_letter = get_column_letter(col_idx)
            ws.conditional_formatting.add(
                f'{col_letter}2:{col_letter}{ws.max_row}',
                ColorScaleRule(
                    start_type='min', start_color='FF00FF00',
                    mid_type='percentile', mid_value=50, mid_color='CC277BF5',
                    end_type='max', end_color='FFFFFF00'
                )
            )
    
    # Add formula column
    feedback_col_idx = None
    base_col_idx = None
    if 'Feedback' in df.columns:
        feedback_col_idx = df.columns.get_loc('Feedback') + 1
    if 'QUESTION ID' in df.columns:
        base_col_idx = df.columns.get_loc('QUESTION ID') + 1
    if feedback_col_idx and base_col_idx:
        feedback_col_letter = get_column_letter(feedback_col_idx)
        base_col_letter = get_column_letter(base_col_idx)
        formula = f'ROUND((COUNTA({feedback_col_letter}:{feedback_col_letter})/COUNTA({base_col_letter}:{base_col_letter}))*100,2)&"%"'
        cell = ws.cell(row=1, column=ws.max_column + 1)
        cell.value = f'={formula}'
        if hasattr(cell, 'data_type'):
            cell.data_type = 'f'
    else:
        ws.cell(row=1, column=ws.max_column + 1).value = ''
    
    # Add end row markers
    end_row = ws.max_row + 1
    for col_name in ['QUESTION ID', 'Feedback']:
        if col_name in df.columns:
            col_idx = df.columns.get_loc(col_name) + 1
            ws.cell(row=end_row, column=col_idx).value = '~{END}~'

    # Freeze the top row and enable autofilter
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    
    # Save to memory buffer and return bytes
    final_buffer = io.BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)
    return final_buffer.getvalue()

# Keep the original function for backwards compatibility
def advanced_process_excel(input_path, output_path):
    
    # Read and drop first 4 rows
    df = pd.read_excel(input_path, skiprows=4)
    
    # Drop empty columns
    df = df.dropna(axis=1, how='all')
    
    # Build final column order (required, added, kept)
    preferred_order = [
        'SURVEY ID', 'ACCOUNT NAME', 'CSM OWNER NAME', 'PM WHO CREATED QUAL', 'EMAIL',
        'DATE FLAGGED', 'QUESTION TYPE', 'QUESTION VISIBILITY', 'DATE CREATED', 'COUNTRY LANGUAGE',
        'QUESTION ID', 'Feedback', 'QUESTION NAME', 'QUESTION TEXT', 'Recommendation',
        'ANSWER PRECODE', 'ANSWER OPTION TEXT'
    ]
    
    # Remove unwanted columns and ensure all preferred columns exist
    df.columns = df.columns.str.strip()
    cols_to_remove = [
        'ACCOUNT ID', 'BUSINESS UNIT ID', 'BUSINESS UNIT NAME', 'SF ACCOUNT ID',
        'REGION', 'CSM OWNER ID', 'CSM EMAIL', 'SURVEY ISSUE NAME', 'DATE/TIME CREATED'
    ]
    df = df.drop(columns=[col for col in cols_to_remove if col in df.columns], errors='ignore')
    for col in preferred_order:
        if col not in df.columns:
            df[col] = ''
    
    # Fill DATE FLAGGED with current system date for all rows
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    if 'DATE FLAGGED' in df.columns:
        df['DATE FLAGGED'] = today_str
    
    # Reindex to final order
    final_cols = [col for col in preferred_order if col in df.columns]
    df = df[final_cols]
    
    # Sort once by all keys present
    sort_keys = [col for col in ['QUESTION ID', 'SURVEY ID', 'ANSWER PRECODE'] if col in df.columns]
    if sort_keys:
        df = df.sort_values(by=sort_keys, kind='stable')
    
    # Detect language for each cell in 'QUESTION TEXT' column and add a new column 'QUESTION TEXT LANGUAGE'
    if 'QUESTION TEXT' in df.columns:
        def detect_lang_safe(text):
            try:
                return detect(str(text)) if pd.notnull(text) and str(text).strip() else ''
            except LangDetectException:
                return ''
        df['QUESTION TEXT LANGUAGE'] = df['QUESTION TEXT'].apply(detect_lang_safe)
    
    # Set Feedback/Recommendation based on language logic
    if all(col in df.columns for col in ['COUNTRY LANGUAGE', 'QUESTION TEXT LANGUAGE', 'Feedback', 'Recommendation']):
        mask_non_english = ~df['COUNTRY LANGUAGE'].astype(str).str.endswith('-ENG')
        mask_qtext_english = df['QUESTION TEXT LANGUAGE'] == 'en'
        # Where COUNTRY LANGUAGE doesn't end with -ENG and QUESTION TEXT is English
        mask_translate = mask_non_english & mask_qtext_english
        df.loc[mask_translate, 'Feedback'] = 'Language Translation Required'
        df.loc[mask_translate, 'Recommendation'] = 'Translate into correct Language'
        # Where COUNTRY LANGUAGE doesn't end with -ENG and QUESTION TEXT is not English
        mask_ok = mask_non_english & ~mask_qtext_english
        df.loc[mask_ok, 'Feedback'] = 'OK'
        df.loc[mask_ok, 'Recommendation'] = 'OK'
    
    # Save to Excel
    df.to_excel(output_path, index=False, engine='openpyxl')
    
    # Post-process with openpyxl
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    
    # Disable word wrap, unmerge, set width
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=False)
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))
    
    # Set custom column widths (only once, for header row)
    col_widths = {
        'SURVEY ID': 9.57, 'ACCOUNT NAME': 9.57, 'CSM OWNER NAME': 9.57,
        'PM WHO CREATED QUAL': 9.57, 'EMAIL': 9.57, 'DATE FLAGGED': 9.57,
        'QUESTION TYPE': 9.57, 'QUESTION VISIBILITY': 4, 'DATE CREATED': 11,
        'ANSWER PRECODE': 5.71, 'QUESTION TEXT': 57.14, 'Feedback': 28.57,
        'Recommendation': 28.57, 'ANSWER OPTION TEXT': 20, 'COUNTRY LANGUAGE': 17.86,
        'QUESTION ID': 8.57
    }
    # Only set widths for columns in the first row (header)
    for cell in ws[1]:
        header = cell.value
        col_letter = get_column_letter(cell.column)
        if header in col_widths:
            ws.column_dimensions[col_letter].width = col_widths[header]
        else:
            ws.column_dimensions[col_letter].width = 20
    
    # Format DATE CREATED to show only date, hide time
    if 'DATE CREATED' in df.columns:
        date_col_idx = df.columns.get_loc('DATE CREATED') + 1
        for row in ws.iter_rows(min_row=2, min_col=date_col_idx, max_col=date_col_idx, max_row=ws.max_row):
            for cell in row:
                cell.number_format = 'yyyy-mm-dd'
    
    # Right align DATE FLAGGED & DATE CREATED
    for col_name in ['DATE FLAGGED', 'DATE CREATED']:
        if col_name in df.columns:
            col_idx = df.columns.get_loc(col_name) + 1
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx, max_row=ws.max_row):
                for cell in row:
                    cell.alignment = Alignment(horizontal='right')
    
    # Conditional formatting for SURVEY ID, QUESTION ID, DATE CREATED
    for col_name in ['SURVEY ID', 'QUESTION ID', 'DATE CREATED']:
        if col_name in df.columns:
            col_idx = df.columns.get_loc(col_name) + 1
            col_letter = get_column_letter(col_idx)
            # Use a 3-color scale for more colorful differentiation
            ws.conditional_formatting.add(
                f'{col_letter}2:{col_letter}{ws.max_row}',
                ColorScaleRule(
                    start_type='min', start_color='FF00FF00',  # Green for min
                    mid_type='percentile', mid_value=50, mid_color='CC277BF5',  # Blue for mid
                    end_type='max', end_color='FFFFFF00'  # Yellow for max
                )
            )
    
    # Add new column at the end with formula in header cell (no header text)
    # Dynamically determine the correct columns for the formula
    # Example: percentage of non-empty Feedback (or any column) over total rows (e.g., 'QUESTION ID')
    feedback_col_idx = None
    base_col_idx = None
    if 'Feedback' in df.columns:
        feedback_col_idx = df.columns.get_loc('Feedback') + 1
    if 'QUESTION ID' in df.columns:
        base_col_idx = df.columns.get_loc('QUESTION ID') + 1
    if feedback_col_idx and base_col_idx:
        feedback_col_letter = get_column_letter(feedback_col_idx)
        base_col_letter = get_column_letter(base_col_idx)
        formula = f'ROUND((COUNTA({feedback_col_letter}:{feedback_col_letter})/COUNTA({base_col_letter}:{base_col_letter}))*100,2)&"%"'
        cell = ws.cell(row=1, column=ws.max_column + 1)
        cell.value = f'={formula}'
        if hasattr(cell, 'data_type'):
            cell.data_type = 'f'  # Explicitly set as formula if possible
    else:
        # Fallback: add a blank column if columns not found
        ws.cell(row=1, column=ws.max_column + 1).value = ''
    
    # Add new row at the end with '~{END}~' for 'QUESTION ID' and 'Feedback' columns
    end_row = ws.max_row + 1
    for col_name in ['QUESTION ID', 'Feedback']:
        if col_name in df.columns:
            col_idx = df.columns.get_loc(col_name) + 1
            ws.cell(row=end_row, column=col_idx).value = '~{END}~'

    # Freeze the top row (header row)
    ws.freeze_panes = 'A2'

    # Enable autofilter for all columns (last step)
    ws.auto_filter.ref = ws.dimensions
    
    wb.save(output_path)