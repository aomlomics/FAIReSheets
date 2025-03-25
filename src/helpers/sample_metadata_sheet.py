"""
Module for creating the sampleMetadata sheet in FAIReSheets.
"""

import pandas as pd
import numpy as np
import time
import json
import gspread_formatting as gsf

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
    
    # Wait to avoid rate limits - reduced from 2 seconds to 1 second
    time.sleep(1)
    
    # Prepare all formatting in a single batch request
    batch_requests = []
    
    # Format header row (term_name row)
    batch_requests.append({
        "repeatCell": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": int(term_name_row),
                "endRowIndex": int(term_name_row) + 1,
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
                    red = float(color_obj.backgroundColor.red)
                    green = float(color_obj.backgroundColor.green)
                    blue = float(color_obj.backgroundColor.blue)
                    
                    batch_requests.append({
                        "repeatCell": {
                            "range": {
                                "sheetId": worksheet.id,
                                "startRowIndex": int(req_level_row),
                                "endRowIndex": int(req_level_row) + 1,
                                "startColumnIndex": int(i),
                                "endColumnIndex": int(i) + 1
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "backgroundColor": {
                                        "red": red,
                                        "green": green,
                                        "blue": blue
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.backgroundColor"
                        }
                    })
    
    # Get column names for reference
    term_names = sheet_df.iloc[term_name_row].tolist()
    
    # Add dropdowns - only for a few rows to reduce API calls
    for i, term in enumerate(term_names):
        if not term or pd.isna(term):
            continue
            
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
                            "startRowIndex": int(term_name_row) + 1,
                            "endRowIndex": int(term_name_row) + 10,  # Only add validation for a few rows
                            "startColumnIndex": int(i),
                            "endColumnIndex": int(i) + 1
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
    
    # Add comments
    for i, term in enumerate(term_names):
        if not term or pd.isna(term) or term in (sampleMetadata_user or []):
            continue
            
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
            
            batch_requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": int(term_name_row),
                        "endRowIndex": int(term_name_row) + 1,
                        "startColumnIndex": int(i),
                        "endColumnIndex": int(i) + 1
                    },
                    "rows": [{"values": [{"note": comment}]}],
                    "fields": "note"
                }
            })
    
    # Apply all formatting in smaller batches to avoid quota limits
    if batch_requests:
        # Split requests into smaller batches - increased from 5 to 8
        batch_size = 8
        for i in range(0, len(batch_requests), batch_size):
            batch_chunk = batch_requests[i:i+batch_size]
            try:
                # Convert any NumPy types to Python native types
                # Serialize and deserialize to convert NumPy types to native Python types
                batch_json = json.dumps({'requests': batch_chunk}, default=lambda x: int(x) if hasattr(x, 'dtype') else float(x) if isinstance(x, (np.float32, np.float64)) else str(x))
                batch_native = json.loads(batch_json)
                
                worksheet.spreadsheet.batch_update(batch_native)
                # Reduced delay between batches from 2 to 1 second
                time.sleep(1)
            except gspread.exceptions.APIError as e:
                if "429" in str(e):  # Rate limit error
                    print(f"Warning: Hit API rate limit at batch {i//batch_size + 1}. Some formatting may not be applied.")
                    # Try a longer delay and continue with fewer requests - reduced from 5 to 3
                    time.sleep(3)
                    try:
                        # Try with even smaller batch
                        for req in batch_chunk:
                            try:
                                # Convert any NumPy types to Python native types
                                req_json = json.dumps({'requests': [req]}, default=lambda x: int(x) if hasattr(x, 'dtype') else float(x) if isinstance(x, (np.float32, np.float64)) else str(x))
                                req_native = json.loads(req_json)
                                
                                worksheet.spreadsheet.batch_update(req_native)
                                time.sleep(0.5)  # Reduced from 1 to 0.5 seconds
                            except Exception as inner_e:
                                print(f"Warning: Error applying individual formatting: {str(inner_e)}")
                    except Exception as batch_e:
                        print(f"Warning: Error in fallback formatting: {str(batch_e)}")
                else:
                    # For other errors, just print a warning and continue
                    print(f"Warning: Error applying formatting: {str(e)}")
            except Exception as e:
                print(f"Warning: Unexpected error: {str(e)}") 