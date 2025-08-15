# This script is for advanced Excel processing using pandas and openpyxl
# It will be imported and used in app.py for the process_excel function
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation
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
        'ANSWER PRECODE', 'ANSWER OPTION TEXT',
        'SF ACCOUNT ID', 'CSM OWNER ID', 'SURVEY ISSUE NAME',
        'Lucid_Action__c', 'OwnerId'  # <-- new columns
    ]
    
    # Remove unwanted columns and ensure all preferred columns exist
    df.columns = df.columns.str.strip()
    cols_to_remove = [
        'ACCOUNT ID', 'BUSINESS UNIT ID', 'BUSINESS UNIT NAME',
        'SF ACCOUNT ID', 'REGION', 'CSM OWNER ID', 'CSM EMAIL', 'DATE/TIME CREATED'
        # 'SURVEY ISSUE NAME' is NOT in remove list
    ]
    # Remove unwanted columns except those to keep
    cols_to_remove = [col for col in cols_to_remove if col not in ['SF ACCOUNT ID', 'CSM OWNER ID', 'SURVEY ISSUE NAME']]
    df = df.drop(columns=[col for col in cols_to_remove if col in df.columns], errors='ignore')
    for col in preferred_order:
        if col not in df.columns:
            df[col] = ''
    
    # Fill DATE FLAGGED with current system date for all rows
    now_str = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
    if 'DATE FLAGGED' in df.columns:
        df['DATE FLAGGED'] = now_str
    
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
        mask_non_english = ~df['COUNTRY LANGUAGE'].astype(str).str.startswith('English')
        mask_qtext_english = df['QUESTION TEXT LANGUAGE'] == 'en'
        # Where COUNTRY LANGUAGE doesn't start with English and QUESTION TEXT is English
        mask_translate = mask_non_english & mask_qtext_english
        df.loc[mask_translate, 'Feedback'] = 'Language Translation Required'
        df.loc[mask_translate, 'Recommendation'] = 'Translate into correct Language'
        # Where COUNTRY LANGUAGE doesn't start with English and QUESTION TEXT is not English
        mask_ok = mask_non_english & ~mask_qtext_english
        df.loc[mask_ok, 'Feedback'] = 'OK'
        df.loc[mask_ok, 'Recommendation'] = 'OK'
    
    # Column rename mapping
    rename_map = {
        'SURVEY ISSUE NAME': 'Name',
        'SURVEY ID': 'Survey_Number__c',
        'PM WHO CREATED QUAL': 'Project_Manager__c',
        'EMAIL': 'Project_Manager_Email__c',
        'DATE FLAGGED': 'Date_Flagged__c',
        'QUESTION TYPE': 'Question_Type__c',
        'QUESTION VISIBILITY': 'Question_Visibility__c',
        'DATE CREATED': 'Custom_Create_Date__c',
        'COUNTRY LANGUAGE': 'Country_Language__c',
        'QUESTION ID': 'QuestionID__c',
        'Feedback': 'Required_Action__c',
        'QUESTION NAME': 'QuestionName__c',
        'QUESTION TEXT': 'Custom_Flagged__c',
        'Recommendation': 'Recommendation__c',
        'SF ACCOUNT ID': 'Buyer_Account__c',
        'CSM OWNER ID': 'CSM_Name__c',  # updated: rename CSM OWNER ID -> CSM_Name__c
        'Lucid_Action__c': 'Setup_Issue__c',
        'ANSWER PRECODE': '[D]ANSPRECODE',
        'ANSWER OPTION TEXT': '[D]OPTION TEXT'
    }
    # Rename columns
    df = df.rename(columns=rename_map)
    df['Lucid_Action__c'] = 'Not Paused'
    # Pre-fill Setup_Issue__c with "--None--" for all rows
    df['Setup_Issue__c'] = '--None--'
    # Define strict output column order
    strict_order = [
        'Name',
        'Buyer_Account__c',
        'CSM_Name__c',
        'Project_Manager__c',
        'Project_Manager_Email__c',
        'Date_Flagged__c',
        'Question_Visibility__c',
        'Custom_Create_Date__c',
        'Lucid_Action__c',
        'OwnerId',
        'Survey_Number__c',
        'Question_Type__c',
        'Country_Language__c',
        'QuestionID__c',
        'Required_Action__c',
        'Recommendation__c',
        'Setup_Issue__c',
        'QuestionName__c',
        'Custom_Flagged__c',
        '[D]ANSPRECODE',
        '[D]OPTION TEXT'
    ]
    # Ensure all columns in strict_order exist
    for col in strict_order:
        if col not in df.columns:
            df[col] = ''
    # Reindex to strict order, drop all others except formula column
    df = df[strict_order]
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

    # Set custom column widths (supports renamed headers)
    orig_col_widths = {
        'SURVEY ID': 9.57, 'ACCOUNT NAME': 9.57, 'CSM OWNER NAME': 9.57,
        'PM WHO CREATED QUAL': 9.57, 'EMAIL': 9.57, 'DATE FLAGGED': 9.57,
        'QUESTION TYPE': 9.57, 'QUESTION VISIBILITY': 4, 'DATE CREATED': 11,
        'ANSWER PRECODE': 5.71, 'QUESTION TEXT': 57.14, 'Feedback': 28.57,
        'Recommendation': 28.57, 'ANSWER OPTION TEXT': 20, 'COUNTRY LANGUAGE': 17.86,
        'QUESTION ID': 8.57
    }
    # build reverse map to find original key for a renamed header
    reverse_rename = {v: k for k, v in rename_map.items()}
    for cell in ws[1]:
        header = cell.value
        col_letter = get_column_letter(cell.column)
        # look up original header name if header was renamed
        orig_header = reverse_rename.get(header, header)
        if orig_header in orig_col_widths:
            ws.column_dimensions[col_letter].width = orig_col_widths[orig_header]
        else:
            ws.column_dimensions[col_letter].width = 20

    # Format DATE CREATED (handle renamed header)
    # If the original 'DATE CREATED' was renamed to 'Custom_Create_Date__c', use df to find column index
    if 'Custom_Create_Date__c' in df.columns:
        date_col_idx = df.columns.get_loc('Custom_Create_Date__c') + 1
        for row in ws.iter_rows(min_row=2, min_col=date_col_idx, max_col=date_col_idx, max_row=ws.max_row):
            for cell in row:
                cell.number_format = 'yyyy-mm-dd'
    
    # Right align DATE FLAGGED & DATE CREATED (use renamed names)
    for col_name in ['Date_Flagged__c', 'Custom_Create_Date__c']:
        if col_name in df.columns:
            col_idx = df.columns.get_loc(col_name) + 1
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx, max_row=ws.max_row):
                for cell in row:
                    cell.alignment = Alignment(horizontal='right')

    # Conditional formatting for renamed columns: Survey_Number__c, QuestionID__c, Custom_Create_Date__c
    for col_name in ['Survey_Number__c', 'QuestionID__c', 'Custom_Create_Date__c']:
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
    
    # Update formula logic to use renamed columns
    feedback_col_idx = None
    base_col_idx = None
    if 'Required_Action__c' in df.columns:
        feedback_col_idx = df.columns.get_loc('Required_Action__c') + 1
    if 'QuestionID__c' in df.columns:
        base_col_idx = df.columns.get_loc('QuestionID__c') + 1
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
    # Remove the end row markers section
    # Add end row markers for renamed columns
    # end_row = ws.max_row + 1
    # for col_name in [
    #     'QuestionID__c',
    #     'Required_Action__c',
    #     'Country_Language__c',
    #     'Recommendation__c',
    #     'Setup_Issue__c'
    # ]:
    #     if col_name in df.columns:
    #         col_idx = df.columns.get_loc(col_name) + 1
    #         ws.cell(row=end_row, column=col_idx).value = '~{END}~'

    # Freeze the top row and enable autofilter
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    
    # Add data validation dropdown for Setup_Issue__c column
    if 'Setup_Issue__c' in df.columns:
        setup_col_idx = df.columns.get_loc('Setup_Issue__c') + 1
        setup_col_letter = get_column_letter(setup_col_idx)
        
        # Define dropdown options with --None-- as first option
        dropdown_options = [
            "--None--",
            "Custom Instead of Standard",
            "Narrow Custom without Targeting", 
            "Leading (Yes/No)",
            "Leading (Too Few Options)",
            "Indicates PII",
            "Multi-Phase + Other Issues",
            "Grammatical Error / Inconsistent Phrasing",
            "Untranslated Text",
            "Illogical Order",
            "Multiple Errors",
            "Other Set up Issue"
        ]
        
        # Create a helper sheet with the dropdown values
        helper_sheet = wb.create_sheet("ValidationData")
        for i, option in enumerate(dropdown_options, 1):
            helper_sheet.cell(row=i, column=1, value=option)
        
        # Add Required_Action__c dropdown options to the same helper sheet
        req_action_options = [
            "Rephrase + Need Additional Options",
            "Custom instead of standard", 
            "Language Translation Required"
        ]
        for i, option in enumerate(req_action_options, 1):
            helper_sheet.cell(row=i, column=2, value=option)
        
        # Create data validation for Setup_Issue__c (strict)
        dv = DataValidation(
            type="list",
            formula1=f"ValidationData!$A$1:$A${len(dropdown_options)}",
            allow_blank=False,
            showErrorMessage=True,
            showInputMessage=True
        )
        dv.error = 'You must select a value from the dropdown list only. Manual typing is not allowed.'
        dv.errorTitle = 'Invalid Entry'
        dv.errorStyle = 'stop'  # This prevents invalid entries
        dv.prompt = 'Please select from the dropdown list only'
        dv.promptTitle = 'Setup Issue Options'
        
        # Apply to all data rows in the Setup_Issue__c column
        ws.add_data_validation(dv)
        dv.add(f'{setup_col_letter}2:{setup_col_letter}{ws.max_row}')
        
        # Create data validation for Required_Action__c (allows custom values)
        if 'Required_Action__c' in df.columns:
            req_action_col_idx = df.columns.get_loc('Required_Action__c') + 1
            req_action_col_letter = get_column_letter(req_action_col_idx)
            
            dv_req = DataValidation(
                type="list",
                formula1=f"ValidationData!$B$1:$B${len(req_action_options)}",
                allow_blank=True,
                showErrorMessage=False,  # Allow custom values
                showInputMessage=True
            )
            dv_req.prompt = 'Select from dropdown or type your own value'
            dv_req.promptTitle = 'Required Action Options'
            
            # Apply to all data rows in the Required_Action__c column
            ws.add_data_validation(dv_req)
            dv_req.add(f'{req_action_col_letter}2:{req_action_col_letter}{ws.max_row}')
        
        # Hide the helper sheet
        helper_sheet.sheet_state = 'hidden'
    
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
        'ANSWER PRECODE', 'ANSWER OPTION TEXT',
        'SF ACCOUNT ID', 'CSM OWNER ID', 'SURVEY ISSUE NAME',
        'Lucid_Action__c', 'OwnerId'  # <-- new columns
    ]
    
    # Remove unwanted columns and ensure all preferred columns exist
    df.columns = df.columns.str.strip()
    cols_to_remove = [
        'ACCOUNT ID', 'BUSINESS UNIT ID', 'BUSINESS UNIT NAME',
        'SF ACCOUNT ID', 'REGION', 'CSM OWNER ID', 'CSM EMAIL', 'DATE/TIME CREATED'
        # 'SURVEY ISSUE NAME' is NOT in remove list
    ]
    # Remove unwanted columns except those to keep
    cols_to_remove = [col for col in cols_to_remove if col not in ['SF ACCOUNT ID', 'CSM OWNER ID', 'SURVEY ISSUE NAME']]
    df = df.drop(columns=[col for col in cols_to_remove if col in df.columns], errors='ignore')
    for col in preferred_order:
        if col not in df.columns:
            df[col] = ''
    
    # Fill DATE FLAGGED with current system date for all rows
    now_str = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
    if 'DATE FLAGGED' in df.columns:
        df['DATE FLAGGED'] = now_str
    
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
        mask_non_english = ~df['COUNTRY LANGUAGE'].astype(str).str.startswith('English')
        mask_qtext_english = df['QUESTION TEXT LANGUAGE'] == 'en'
        # Where COUNTRY LANGUAGE doesn't start with English and QUESTION TEXT is English
        mask_translate = mask_non_english & mask_qtext_english
        df.loc[mask_translate, 'Feedback'] = 'Language Translation Required'
        df.loc[mask_translate, 'Recommendation'] = 'Translate into correct Language'
        # Where COUNTRY LANGUAGE doesn't start with English and QUESTION TEXT is not English
        mask_ok = mask_non_english & ~mask_qtext_english
        df.loc[mask_ok, 'Feedback'] = 'OK'
        df.loc[mask_ok, 'Recommendation'] = 'OK'
    
    # Column rename mapping
    rename_map = {
        'SURVEY ISSUE NAME': 'Name',
        'SURVEY ID': 'Survey_Number__c',
        'PM WHO CREATED QUAL': 'Project_Manager__c',
        'EMAIL': 'Project_Manager_Email__c',
        'DATE FLAGGED': 'Date_Flagged__c',
        'QUESTION TYPE': 'Question_Type__c',
        'QUESTION VISIBILITY': 'Question_Visibility__c',
        'DATE CREATED': 'Custom_Create_Date__c',
        'COUNTRY LANGUAGE': 'Country_Language__c',
        'QUESTION ID': 'QuestionID__c',
        'Feedback': 'Required_Action__c',
        'QUESTION NAME': 'QuestionName__c',
        'QUESTION TEXT': 'Custom_Flagged__c',
        'Recommendation': 'Recommendation__c',
        'SF ACCOUNT ID': 'Buyer_Account__c',
        'CSM OWNER ID': 'CSM_Name__c',  # updated: rename CSM OWNER ID -> CSM_Name__c
        'Lucid_Action__c': 'Setup_Issue__c',
        'ANSWER PRECODE': '[D]ANSPRECODE',
        'ANSWER OPTION TEXT': '[D]OPTION TEXT'
    }
    # Rename columns
    df = df.rename(columns=rename_map)
    df['Lucid_Action__c'] = 'Not Paused'
    # Pre-fill Setup_Issue__c with "--None--" for all rows
    df['Setup_Issue__c'] = '--None--'
    strict_order = [
        'Name',
        'Buyer_Account__c',
        'CSM_Name__c',
        'Project_Manager__c',
        'Project_Manager_Email__c',
        'Date_Flagged__c',
        'Question_Visibility__c',
        'Custom_Create_Date__c',
        'Lucid_Action__c',
        'OwnerId',
        'Survey_Number__c',
        'Question_Type__c',
        'Country_Language__c',
        'QuestionID__c',
        'Required_Action__c',
        'Recommendation__c',
        'Setup_Issue__c',
        'QuestionName__c',
        'Custom_Flagged__c',
        '[D]ANSPRECODE',
        '[D]OPTION TEXT'
    ]
    for col in strict_order:
        if col not in df.columns:
            df[col] = ''
    df = df[strict_order]
    df.to_excel(output_path, index=False, engine='openpyxl')
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active

    # Disable word wrap, unmerge, set width
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=False)
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))

    # Set custom column widths (supports renamed headers)
    orig_col_widths = {
        'SURVEY ID': 9.57, 'ACCOUNT NAME': 9.57, 'CSM OWNER NAME': 9.57,
        'PM WHO CREATED QUAL': 9.57, 'EMAIL': 9.57, 'DATE FLAGGED': 9.57,
        'QUESTION TYPE': 9.57, 'QUESTION VISIBILITY': 4, 'DATE CREATED': 11,
        'ANSWER PRECODE': 5.71, 'QUESTION TEXT': 57.14, 'Feedback': 28.57,
        'Recommendation': 28.57, 'ANSWER OPTION TEXT': 20, 'COUNTRY LANGUAGE': 17.86,
        'QUESTION ID': 8.57
    }
    reverse_rename = {v: k for k, v in rename_map.items()}
    for cell in ws[1]:
        header = cell.value
        col_letter = get_column_letter(cell.column)
        orig_header = reverse_rename.get(header, header)
        if orig_header in orig_col_widths:
            ws.column_dimensions[col_letter].width = orig_col_widths[orig_header]
        else:
            ws.column_dimensions[col_letter].width = 20

    # Format DATE CREATED (handle renamed header)
    if 'Custom_Create_Date__c' in df.columns:
        date_col_idx = df.columns.get_loc('Custom_Create_Date__c') + 1
        for row in ws.iter_rows(min_row=2, min_col=date_col_idx, max_col=date_col_idx, max_row=ws.max_row):
            for cell in row:
                cell.number_format = 'yyyy-mm-dd'
    
    # Right align DATE FLAGGED & DATE CREATED (use renamed names)
    for col_name in ['Date_Flagged__c', 'Custom_Create_Date__c']:
        if col_name in df.columns:
            col_idx = df.columns.get_loc(col_name) + 1
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx, max_row=ws.max_row):
                for cell in row:
                    cell.alignment = Alignment(horizontal='right')

    # Conditional formatting for renamed columns: Survey_Number__c, QuestionID__c, Custom_Create_Date__c
    for col_name in ['Survey_Number__c', 'QuestionID__c', 'Custom_Create_Date__c']:
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

    # Update formula logic to use renamed columns
    feedback_col_idx = None
    base_col_idx = None
    if 'Required_Action__c' in df.columns:
        feedback_col_idx = df.columns.get_loc('Required_Action__c') + 1
    if 'QuestionID__c' in df.columns:
        base_col_idx = df.columns.get_loc('QuestionID__c') + 1
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
    # Remove the end row markers section
    # end_row = ws.max_row + 1
    # for col_name in [
    #     'QuestionID__c',
    #     'Required_Action__c',
    #     'Country_Language__c',
    #     'Recommendation__c',
    #     'Setup_Issue__c'
    # ]:
    #     if col_name in df.columns:
    #         col_idx = df.columns.get_loc(col_name) + 1
    #         ws.cell(row=end_row, column=col_idx).value = '~{END}~'

    # Freeze the top row (header row)
    ws.freeze_panes = 'A2'

    # Enable autofilter for all columns (last step)
    ws.auto_filter.ref = ws.dimensions
    
    # Add data validation dropdown for Setup_Issue__c column
    if 'Setup_Issue__c' in df.columns:
        setup_col_idx = df.columns.get_loc('Setup_Issue__c') + 1
        setup_col_letter = get_column_letter(setup_col_idx)
        
        # Define dropdown options with --None-- as first option
        dropdown_options = [
            "--None--",
            "Custom Instead of Standard",
            "Narrow Custom without Targeting", 
            "Leading (Yes/No)",
            "Leading (Too Few Options)",
            "Indicates PII",
            "Multi-Phase + Other Issues",
            "Grammatical Error / Inconsistent Phrasing",
            "Untranslated Text",
            "Illogical Order",
            "Multiple Errors",
            "Other Set up Issue"
        ]
        
        # Create a helper sheet with the dropdown values
        helper_sheet = wb.create_sheet("ValidationData")
        for i, option in enumerate(dropdown_options, 1):
            helper_sheet.cell(row=i, column=1, value=option)
        
        # Add Required_Action__c dropdown options to the same helper sheet
        req_action_options = [
            "Rephrase + Need Additional Options",
            "Custom instead of standard", 
            "Language Translation Required"
        ]
        for i, option in enumerate(req_action_options, 1):
            helper_sheet.cell(row=i, column=2, value=option)
        
        # Create data validation for Setup_Issue__c (strict)
        dv = DataValidation(
            type="list",
            formula1=f"ValidationData!$A$1:$A${len(dropdown_options)}",
            allow_blank=False,
            showErrorMessage=True,
            showInputMessage=True
        )
        dv.error = 'You must select a value from the dropdown list only. Manual typing is not allowed.'
        dv.errorTitle = 'Invalid Entry'
        dv.errorStyle = 'stop'  # This prevents invalid entries
        dv.prompt = 'Please select from the dropdown list only'
        dv.promptTitle = 'Setup Issue Options'
        
        # Apply to all data rows in the Setup_Issue__c column
        ws.add_data_validation(dv)
        dv.add(f'{setup_col_letter}2:{setup_col_letter}{ws.max_row}')
        
        # Create data validation for Required_Action__c (allows custom values)
        if 'Required_Action__c' in df.columns:
            req_action_col_idx = df.columns.get_loc('Required_Action__c') + 1
            req_action_col_letter = get_column_letter(req_action_col_idx)
            
            dv_req = DataValidation(
                type="list",
                formula1=f"ValidationData!$B$1:$B${len(req_action_options)}",
                allow_blank=True,
                showErrorMessage=False,  # Allow custom values
                showInputMessage=True
            )
            dv_req.prompt = 'Select from dropdown or type your own value'
            dv_req.promptTitle = 'Required Action Options'
            
            # Apply to all data rows in the Required_Action__c column
            ws.add_data_validation(dv_req)
            dv_req.add(f'{req_action_col_letter}2:{req_action_col_letter}{ws.max_row}')
        
        # Hide the helper sheet
        helper_sheet.sheet_state = 'hidden'
    
    wb.save(output_path)