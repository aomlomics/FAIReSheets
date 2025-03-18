"""
Module for creating the experimentRunMetadata sheet in FAIReSheets.
"""

import pandas as pd
import numpy as np
import time
import json
import gspread_formatting as gsf
import gspread

def create_experiment_metadata_sheet(worksheet, full_temp_file_name, input_df, req_lev, color_styles, vocab_df):
    """Create and format the experimentRunMetadata sheet."""
    
    # Read the template using pandas
    sheet_df = pd.read_excel(full_temp_file_name, sheet_name="experimentRunMetadata", header=None)
    
    # Replace NaN values with empty strings to avoid JSON errors
    sheet_df = sheet_df.fillna('')
        
    # Find key rows
    term_name_row = sheet_df.shape[0] - 1  # Last row typically contains term names
        
    # Find the requirement level row
    req_lev_row = None
    section_row = None
    for idx, row in sheet_df.iterrows():
        if '# requirement_level_code' in row.values:
            req_lev_row = idx
        if '# section' in row.values:
            section_row = idx
        
    # Filter by requirement level
    if req_lev_row is not None:
        req_lev2rm = [level for level in ['M', 'HR', 'R', 'O'] if level not in req_lev]
        cols_to_drop = []
        
        for col_idx in range(sheet_df.shape[1]):
            if sheet_df.iloc[req_lev_row, col_idx] in req_lev2rm:
                cols_to_drop.append(col_idx)
        
        # Drop columns with unwanted requirement levels
        if cols_to_drop:
            sheet_df = sheet_df.drop(columns=sheet_df.columns[cols_to_drop])
    
    # Convert to list of lists for gspread
    data = sheet_df.values.tolist()
    
    # Resize worksheet to accommodate all data (add some buffer)
    rows_needed = len(data) + 20  # Add buffer
    cols_needed = len(data[0]) + 10 if data else 50  # Add buffer
    worksheet.resize(rows=rows_needed, cols=cols_needed)
    
    # Update the worksheet with all data at once
    worksheet.update("A1", data)
    
    # Wait a moment to avoid rate limits
    time.sleep(1)
        
    # Prepare batch requests for formatting
    batch_requests = []
    
    # Format term_name row with bold
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
    if req_lev_row is not None:
        for col_idx, req_level in enumerate(data[req_lev_row]):
            if req_level in color_styles and req_level in req_lev:
                color_obj = color_styles[req_level]
                if hasattr(color_obj, 'backgroundColor') and hasattr(color_obj.backgroundColor, 'red'):
                    red = float(color_obj.backgroundColor.red)
                    green = float(color_obj.backgroundColor.green)
                    blue = float(color_obj.backgroundColor.blue)
                    
                    batch_requests.append({
                        "repeatCell": {
                            "range": {
                                "sheetId": worksheet.id,
                                "startRowIndex": req_lev_row,
                                "endRowIndex": req_lev_row + 1,
                                "startColumnIndex": col_idx,
                                "endColumnIndex": col_idx + 1
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
    
    # Get column names from the term_name row
    term_names = data[term_name_row] if term_name_row < len(data) else []
        
    # Add dropdowns for vocabulary fields
    for col_idx, term in enumerate(term_names):
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
                            "startRowIndex": term_name_row + 1,
                            "endRowIndex": term_name_row + 20,  # Add validation for several rows
                            "startColumnIndex": col_idx,
                            "endColumnIndex": col_idx + 1
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
    
    # Add comments with field information
    for col_idx, term in enumerate(term_names):
        if not term or pd.isna(term):
            continue
        
        term_info = input_df[input_df['term_name'] == term]
        if not term_info.empty:
            # Build comment text
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
                        "startRowIndex": term_name_row,
                        "endRowIndex": term_name_row + 1,
                        "startColumnIndex": col_idx,
                        "endColumnIndex": col_idx + 1
                    },
                    "rows": [{"values": [{"note": comment}]}],
                    "fields": "note"
                }
            })
    
    # Apply all formatting in smaller batches to avoid quota limits
    if batch_requests:
        # Split requests into smaller batches
        batch_size = 5  # Process 5 requests at a time
        for i in range(0, len(batch_requests), batch_size):
            batch_chunk = batch_requests[i:i+batch_size]
            try:
                # Convert any NumPy types to Python native types
                # Serialize and deserialize to convert NumPy types to native Python types
                batch_json = json.dumps({'requests': batch_chunk}, default=lambda x: int(x) if hasattr(x, 'dtype') else float(x) if isinstance(x, (np.float32, np.float64)) else str(x))
                batch_native = json.loads(batch_json)
                
                worksheet.spreadsheet.batch_update(batch_native)
                # Add delay between batches
                time.sleep(1)
            except gspread.exceptions.APIError as e:
                if "429" in str(e):  # Rate limit error
                    print(f"Warning: Hit API rate limit at batch {i//batch_size + 1}. Some formatting may not be applied.")
                    # Try a longer delay and continue with fewer requests
                    time.sleep(5)
                    try:
                        # Try with even smaller batch
                        for req in batch_chunk:
                            try:
                                # Convert any NumPy types to Python native types
                                req_json = json.dumps({'requests': [req]}, default=lambda x: int(x) if hasattr(x, 'dtype') else float(x) if isinstance(x, (np.float32, np.float64)) else str(x))
                                req_native = json.loads(req_json)
                                
                                worksheet.spreadsheet.batch_update(req_native)
                                time.sleep(1)
                            except Exception as inner_e:
                                print(f"Warning: Error applying individual formatting: {str(inner_e)}")
                    except Exception as batch_e:
                        print(f"Warning: Error in fallback formatting: {str(batch_e)}")
                else:
                    # For other errors, just print a warning and continue
                    print(f"Warning: Error applying formatting: {str(e)}")
            except Exception as e:
                print(f"Warning: Unexpected error: {str(e)}") 