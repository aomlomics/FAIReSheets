"""
Module for creating the experimentRunMetadata sheet in FAIReSheets.
"""

import pandas as pd
import numpy as np
import time
import json
import gspread_formatting as gsf
import gspread

def create_experiment_metadata_sheet(worksheet, full_temp_file_name, input_df, req_lev, color_styles, vocab_df, experimentRunMetadata_user=None):
    """Create and format the experimentRunMetadata sheet."""
    
    # Read the template using pandas
    sheet_df = pd.read_excel(full_temp_file_name, sheet_name="experimentRunMetadata", header=None)
    
    # Replace NaN values with empty strings to avoid JSON errors
    sheet_df = sheet_df.fillna('')
        
    # Find key rows
    term_name_row = None
    req_lev_row = None
    section_row = None
    
    for idx, row in sheet_df.iterrows():
        if '# requirement_level_code' in row.values:
            req_lev_row = idx
        if '# section' in row.values:
            section_row = idx
        # The term name row is typically the last row with actual field names
        if any(col for col in row if isinstance(col, str) and not col.startswith('#')):
            term_name_row = idx
        
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
    
    # Add user-defined fields if provided
    if experimentRunMetadata_user and term_name_row is not None and req_lev_row is not None and section_row is not None:
        for field in experimentRunMetadata_user:
            # Create a new column for the user field
            new_col_idx = sheet_df.shape[1]
            sheet_df[new_col_idx] = ''
            
            # Set the field name, requirement level, and section
            sheet_df.iloc[term_name_row, new_col_idx] = field
            sheet_df.iloc[req_lev_row, new_col_idx] = 'O'  # Optional
            sheet_df.iloc[section_row, new_col_idx] = 'User defined'
    
    # Convert to list of lists for gspread
    data = sheet_df.values.tolist()
    
    # Resize worksheet to accommodate all data (add some buffer)
    rows_needed = len(data) + 20  # Add buffer
    cols_needed = len(data[0]) + 10 if data else 50  # Add buffer
    worksheet.resize(rows=rows_needed, cols=cols_needed)
    
    # Update the worksheet with all data at once
    worksheet.update("A1", data)
    
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
                                        "red": float(color_obj.backgroundColor.red),
                                        "green": float(color_obj.backgroundColor.green),
                                        "blue": float(color_obj.backgroundColor.blue)
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.backgroundColor"
                        }
                    })
    
    # Get column names from the term_name row
    term_names = data[term_name_row] if term_name_row < len(data) else []
        
    # Add dropdowns and comments in batches
    for col_idx, term in enumerate(term_names):
        if not term or pd.isna(term):
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
                            "endRowIndex": term_name_row + 20,
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
        
        # Handle comments
        term_info = input_df[input_df['term_name'] == term]
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
                        "startRowIndex": term_name_row,
                        "endRowIndex": term_name_row + 1,
                        "startColumnIndex": col_idx,
                        "endColumnIndex": col_idx + 1
                    },
                    "rows": [{"values": [{"note": comment}]}],
                    "fields": "note"
                }
            })
    
    # Apply all formatting, dropdowns, and notes in one batch
    if batch_requests:
        try:
            worksheet.spreadsheet.batch_update({'requests': batch_requests})
        except gspread.exceptions.APIError as e:
            if "429" in str(e):  # Rate limit error
                print("Warning: Hit API rate limit. Waiting 60 seconds before retrying...")
                time.sleep(60)  # Wait a full minute
                worksheet.spreadsheet.batch_update({'requests': batch_requests})
            else:
                raise
    
    # Reduced wait time (from 1 second to 0.5)
    time.sleep(0.5) 