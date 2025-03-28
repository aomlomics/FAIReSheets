"""
Module for creating the sampleMetadata sheet in FAIReSheets.
"""

import pandas as pd
import numpy as np
import time
import json
import gspread_formatting as gsf
import gspread

def create_sample_metadata_sheet(worksheet, full_temp_file_name, input_df, req_lev, sample_type,
                                 assay_type, assay_name, sampleMetadata_user, color_styles, vocab_df):
    """Create and format the sampleMetadata sheet."""
    
    # Read the template sheet
    sheet_df = pd.read_excel(full_temp_file_name, sheet_name="sampleMetadata", header=None)
    sheet_df = sheet_df.fillna('')
    
    # Find key rows
    term_name_row = sheet_df[sheet_df.iloc[:, 0] == 'samp_name'].index[0]
    section_row = sheet_df[sheet_df.iloc[:, 0] == '# section'].index[0]
    req_level_row = sheet_df[sheet_df.iloc[:, 0] == '# requirement_level_code'].index[0]
    
    # Convert NumPy integers to Python integers to avoid JSON serialization issues
    term_name_row = int(term_name_row)
    section_row = int(section_row)
    req_level_row = int(req_level_row)
    
    # Temporarily set column names for easier filtering
    temp_col_names = sheet_df.iloc[term_name_row].tolist()
    sheet_df.columns = [str(col) if pd.notna(col) and col != '' else f"col_{i}" 
                        for i, col in enumerate(temp_col_names)]
    
    # Filter by sample type if not 'other'
    if not any(s.lower() == 'other' for s in sample_type):
        filtered_terms = input_df[
            input_df['sample_type_specificity'].isna() | 
            (input_df['sample_type_specificity'] == 'ALL') |
            input_df['sample_type_specificity'].str.contains('|'.join(sample_type), case=False, na=False)
        ]['term_name'].tolist()
        
        # Make sure to keep the first column (with samp_name)
        cols_to_keep = ['col_0'] if 'col_0' in sheet_df.columns else []
        
        # Add other columns that match filtered terms or are empty/unnamed
        cols_to_keep.extend([col for col in sheet_df.columns 
                           if col in filtered_terms or (col.startswith('col_') and col != 'col_0')])
        
        sheet_df = sheet_df[cols_to_keep]
    
    # Remove 'Targeted assay detection' section for metabarcoding
    if assay_type == 'metabarcoding':
        section_values = sheet_df.iloc[section_row].tolist()
        cols_to_drop = [col for i, col in enumerate(sheet_df.columns) 
                       if i < len(section_values) and section_values[i] == 'Targeted assay detection']
        if cols_to_drop:
            sheet_df = sheet_df.drop(columns=cols_to_drop)
    
    # Handle detected_notDetected for targeted assays with multiple assay names
    if assay_type == 'targeted' and len(assay_name) > 1:
        # Find detected_notDetected column
        detected_col = next((col for col in sheet_df.columns 
                           if col == 'detected_notDetected'), None)
        
        if detected_col:
            # Get the requirement level
            req_level_value = sheet_df.iloc[req_level_row][detected_col]
            
            # Drop the original column
            sheet_df = sheet_df.drop(columns=[detected_col])
            
            # Add new columns for each assay
            for name in assay_name:
                col_name = f'detected_notDetected_{name}'
                sheet_df[col_name] = ''
                sheet_df.iloc[term_name_row, sheet_df.columns.get_loc(col_name)] = col_name
                sheet_df.iloc[req_level_row, sheet_df.columns.get_loc(col_name)] = req_level_value
                sheet_df.iloc[section_row, sheet_df.columns.get_loc(col_name)] = 'Targeted assay detection'
    
    # Filter by requirement level
    req_levels = sheet_df.iloc[req_level_row].tolist()
    cols_to_keep = [0]  # Always keep the first column
    cols_to_keep.extend([i for i, level in enumerate(req_levels) 
                       if (i > 0) and (level in req_lev or pd.isna(level) or level == '')])
    sheet_df = sheet_df.iloc[:, cols_to_keep]
    
    # Add user-defined fields
    if sampleMetadata_user:
        for field in sampleMetadata_user:
            sheet_df[field] = ''
            sheet_df.iloc[term_name_row, sheet_df.columns.get_loc(field)] = field
            sheet_df.iloc[req_level_row, sheet_df.columns.get_loc(field)] = 'O'
            sheet_df.iloc[section_row, sheet_df.columns.get_loc(field)] = 'User defined'
    
    # Pre-fill assay_name if it exists
    assay_name_col = next((col for col in sheet_df.columns if col == 'assay_name'), None)
    if assay_name_col:
        # Add just one empty row for data entry
        new_row_idx = term_name_row + 1
        if new_row_idx >= len(sheet_df):
            sheet_df.loc[new_row_idx] = ''
        else:
            sheet_df.iloc[new_row_idx] = ''
        
        # Fill in the assay_name
        sheet_df.iloc[new_row_idx, sheet_df.columns.get_loc(assay_name_col)] = ' | '.join(assay_name)
    else:
        # Add just one empty row for data entry if it doesn't exist already
        new_row_idx = term_name_row + 1
        if new_row_idx >= len(sheet_df):
            sheet_df.loc[new_row_idx] = ''
    
    # Convert to list for Google Sheets
    data = sheet_df.values.tolist()
    
    # Update the worksheet with all data at once - only add a few rows for the user to fill in
    worksheet.resize(rows=int(term_name_row + 10), cols=int(len(data[0]) + 5))  # Just a few rows needed
    worksheet.update("A1", data)
    
    # Prepare batch requests for formatting
    batch_requests = []
    
    # Format header row (term_name row)
    batch_requests.append({
        "repeatCell": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": term_name_row,
                "endRowIndex": term_name_row + 1,
                "startColumnIndex": 0,
                "endColumnIndex": len(data[0])
            },
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {
                        "bold": True
                    }
                }
            },
            "fields": "userEnteredFormat.textFormat.bold"
        }
    })
    
    # Format requirement level cells with colors
    for i, level in enumerate(sheet_df.iloc[req_level_row]):
        if level in color_styles and level in req_lev:
            # Get color from the color_styles dictionary
            color_obj = color_styles[level]
            if hasattr(color_obj, 'backgroundColor'):
                # Extract RGB values from the color object
                if hasattr(color_obj.backgroundColor, 'red'):
                    batch_requests.append({
                        "repeatCell": {
                            "range": {
                                "sheetId": worksheet.id,
                                "startRowIndex": req_level_row,
                                "endRowIndex": req_level_row + 1,
                                "startColumnIndex": i,
                                "endColumnIndex": i + 1
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "backgroundColor": {
                                        "red": float(color_obj.backgroundColor.red),
                                        "green": float(color_obj.backgroundColor.green),
                                        "blue": float(color_obj.backgroundColor.blue)
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.backgroundColor"
                        }
                    })
    
    # Get column names for reference
    term_names = sheet_df.iloc[term_name_row].tolist()
    
    # Prepare note updates list
    note_updates = []
    
    # Add dropdowns and comments in batches
    for i, term in enumerate(term_names):
        if not term or pd.isna(term) or term in (sampleMetadata_user or []):
            continue
            
        # Handle dropdowns
        vocab_row = vocab_df[vocab_df['term_name'] == term]
        if not vocab_row.empty and 'n_options' in vocab_row.columns:
            n_options = int(vocab_row.iloc[0]['n_options'])
            values = [str(vocab_row.iloc[0][f'vocab{j+1}']) for j in range(n_options) 
                    if f'vocab{j+1}' in vocab_row.columns and pd.notna(vocab_row.iloc[0][f'vocab{j+1}'])]
            
            if values:
                batch_requests.append({
                    "setDataValidation": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": term_name_row + 1,
                            "endRowIndex": term_name_row + 10,
                            "startColumnIndex": i,
                            "endColumnIndex": i + 1
                        },
                        "rule": {
                            "condition": {
                                "type": "ONE_OF_LIST",
                                "values": [{"userEnteredValue": v} for v in values]
                            },
                            "showCustomUi": True
                        }
                    }
                })
        
        # Handle comments
        term_for_lookup = 'detected_notDetected' if term.startswith('detected_notDetected_') else term
        term_info = input_df[input_df['term_name'] == term_for_lookup]
        
        if not term_info.empty:
            comment = f"Requirement level: {term_info.iloc[0]['requirement_level']}"
            if not pd.isna(term_info.iloc[0]['requirement_level_condition']):
                comment += f" ({term_info.iloc[0]['requirement_level_condition']})"
            comment += f"\nDescription: {term_info.iloc[0]['description']}"
            comment += f"\nExample: {term_info.iloc[0]['example']}"
            comment += f"\nField type: {term_info.iloc[0]['term_type']}"
            
            if term_info.iloc[0]['term_type'] == 'controlled vocabulary':
                comment += f" ({term_info.iloc[0]['controlled_vocabulary_options']})"
            elif term_info.iloc[0]['term_type'] == 'fixed format':
                comment += f" ({term_info.iloc[0]['fixed_format']})"
            
            note_updates.append({
                "range": f"{chr(65+i)}{term_name_row+1}",
                "note": comment
            })
    
    # Apply all notes in one batch if supported by gspread
    if note_updates:
        # Add note requests to batch_requests
        for note in note_updates:
            batch_requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": term_name_row,
                        "endRowIndex": term_name_row + 1,
                        "startColumnIndex": ord(note["range"][0]) - 65,  # Convert A to 0, B to 1, etc.
                        "endColumnIndex": ord(note["range"][0]) - 64
                    },
                    "rows": [{"values": [{"note": note["note"]}]}],
                    "fields": "note"
                }
            })
        
        # Apply all formatting and notes in one batch
        try:
            worksheet.spreadsheet.batch_update({'requests': batch_requests})
        except gspread.exceptions.APIError as e:
            if "429" in str(e):  # Rate limit error
                print("Warning: Hit API rate limit. Waiting 60 seconds before retrying...")
                time.sleep(60)  # Wait a full minute
                worksheet.spreadsheet.batch_update({'requests': batch_requests})
            else:
                raise
    
    # Wait to avoid rate limits - reduced from 2 seconds to 1 second
    time.sleep(1) 