"""
Module for creating targeted assay sheets in FAIReSheets.

This module contains functions to create sheets for targeted assay data including
standard curves, quantitative data, and amplification data.
"""

import pandas as pd
import numpy as np
import time
import json
import gspread_formatting as gsf
import gspread

def create_targeted_sheets(worksheets, sheet_names, full_temp_file_path, full_template_df, input_df, req_lev, 
                          color_styles, vocab_df, project_id, assay_name):
    """Create and format targeted assay sheets."""
    
    # Process each sheet
    for sheet_name in sheet_names:
        try:
            print(f"\nProcessing {sheet_name} sheet...")
            
            # Read the template directly from the Excel file
            sheet_df = pd.read_excel(full_temp_file_path, sheet_name=sheet_name, header=None)
            print(f"Successfully read sheet from file: {sheet_name}")
            
            # Replace NaN values with empty strings to avoid JSON errors
            sheet_df = sheet_df.fillna('')
            
            print(f"Sheet shape: {sheet_df.shape}")
            
            # For targeted sheets, first two rows are usually:
            # Row 0: # requirement_level_code M M M...
            # Row 1: # section PCR PCR...
            # The actual data columns/fields are in the last row
            
            # Find rows with markers (similar to taxa_sheets.py)
            req_level_row = None
            section_row = None
            for idx, row in sheet_df.iterrows():
                if isinstance(row[0], str) and '# requirement_level_code' in row[0]:
                    req_level_row = idx
                    print(f"Found requirement_level_code at row {idx}")
                if isinstance(row[0], str) and '# section' in row[0]:
                    section_row = idx
                    print(f"Found section at row {idx}")
            
            # The term_name row is the last row (with column headers)
            term_name_row = sheet_df.shape[0] - 1  # Last row contains the actual field names
            print(f"Using term_name_row = {term_name_row}")
            
            # For targeted sheets, we don't have a separate description row in the file
            # Instead, descriptions will be added as notes to cells in Google Sheets
            description_row = None
            
            # Filter columns based on requirement level
            if req_level_row is not None:
                req_lev2rm = [level for level in ['M', 'HR', 'R', 'O'] if level not in req_lev]
                cols_to_drop = []
                
                for col_idx in range(sheet_df.shape[1]):
                    if sheet_df.iloc[req_level_row, col_idx] in req_lev2rm:
                        cols_to_drop.append(col_idx)
                
                # Drop columns with unwanted requirement levels
                if cols_to_drop:
                    sheet_df = sheet_df.drop(columns=sheet_df.columns[cols_to_drop])
                    print(f"Dropped {len(cols_to_drop)} columns with requirement levels not in {req_lev}")
            
            # Convert to list of lists for gspread
            data = sheet_df.values.tolist()
            
            # Resize worksheet to accommodate all data
            rows_needed = len(data) + 20  # Add buffer
            cols_needed = len(data[0]) + 10 if data else 50  # Add buffer
            worksheet = worksheets[sheet_name]
            worksheet.resize(rows=rows_needed, cols=cols_needed)
            
            # Update the worksheet with all data at once
            worksheet.update("A1", data)
            
            # Short delay to avoid rate limits
            time.sleep(0.5)
            
            # Get term names from the last row (these are column names)
            term_names = []
            if term_name_row < len(data):
                term_names = data[term_name_row]
                print(f"Found term names: {term_names[:5]}...")
            else:
                print(f"Warning: term_name_row {term_name_row} is outside data range {len(data)}")
            
            # Add project_id and assay_name if columns exist
            for col_idx, term in enumerate(term_names):
                if term == 'projectID' and project_id:
                    # Update the worksheet cell directly
                    worksheet.update_cell(term_name_row + 2, col_idx + 1, project_id)
                    print(f"Added project_id '{project_id}' to column {col_idx+1}")
                
                if term == 'assayName' and assay_name and len(assay_name) > 0:
                    # Update the worksheet cell directly
                    worksheet.update_cell(term_name_row + 2, col_idx + 1, assay_name[0])
                    print(f"Added assay_name '{assay_name[0]}' to column {col_idx+1}")
            
            # Prepare batch requests for formatting
            batch_requests = []
            
            # Format term_name row with bold (same as taxa_sheets)
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
            if req_level_row is not None:
                for col_idx, req_level in enumerate(data[req_level_row]):
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
                                        "startRowIndex": req_level_row,
                                        "endRowIndex": req_level_row + 1,
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
                batch_size = 8
                for i in range(0, len(batch_requests), batch_size):
                    batch_chunk = batch_requests[i:i+batch_size]
                    try:
                        # Convert any NumPy types to Python native types
                        batch_json = json.dumps({'requests': batch_chunk}, default=lambda x: int(x) if hasattr(x, 'dtype') else float(x) if isinstance(x, (np.float32, np.float64)) else str(x))
                        batch_native = json.loads(batch_json)
                        
                        worksheet.spreadsheet.batch_update(batch_native)
                        time.sleep(0.5)  # Short delay to avoid rate limits
                    except Exception as e:
                        print(f"Warning: Unexpected error in batch update: {str(e)}")
            
            print(f"Successfully completed {sheet_name} sheet")
                
        except Exception as e:
            print(f"Error processing {sheet_name} sheet: {e}")
            import traceback
            traceback.print_exc()
            continue
