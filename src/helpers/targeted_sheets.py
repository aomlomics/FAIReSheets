"""
Module for creating targeted assay sheets in FAIReSheets.
This includes stdData, eLowQuantData, and ampData sheets.
"""

import pandas as pd
import numpy as np
import time
import gspread_formatting as gsf
import json
import gspread

def create_targeted_sheets(worksheets, sheet_names, full_temp_file_name, input_df, req_lev, 
                          color_styles, vocab_df, project_id, assay_name):
    """
    Create and format sheets specific to targeted assays.
    
    Parameters:
    -----------
    worksheets : dict
        Dictionary of worksheet objects
    sheet_names : list
        List of sheet names to create
    full_temp_file_name : str
        Path to the full template file
    input_df : DataFrame
        Input checklist dataframe
    req_lev : list
        Requirement levels to include
    color_styles : dict
        Dictionary of color styles for requirement levels
    vocab_df : DataFrame
        Vocabulary dataframe for dropdowns
    project_id : str
        Project identifier
    assay_name : list
        List of assay names
    """
    # Print the column names of vocab_df for debugging
    print(f"Vocabulary DataFrame columns: {vocab_df.columns.tolist()}")
    
    for sheet_name in sheet_names:
        worksheet = worksheets[sheet_name]
        print(f"Processing {sheet_name} sheet...")
        
        # Read the template using pandas
        try:
            sheet_df = pd.read_excel(full_temp_file_name, sheet_name=sheet_name, header=None)
        except Exception as e:
            print(f"Error reading {sheet_name} from template: {e}")
            continue
        
        # Replace NaN values with empty strings to avoid JSON errors
        sheet_df = sheet_df.fillna('')
        
        # Find key rows if they exist
        term_name_row = None
        req_level_row = None
        section_row = None
        description_row = None
        
        for idx, row in sheet_df.iterrows():
            if '# requirement_level_code' in row.values:
                req_level_row = idx
            if '# section' in row.values:
                section_row = idx
            if '# description' in row.values:
                description_row = idx
            # The term name row is typically the last row with actual field names
            if any(col for col in row if isinstance(col, str) and not col.startswith('#')):
                term_name_row = idx
        
        print(f"Found rows - term_name: {term_name_row}, req_level: {req_level_row}, section: {section_row}, description: {description_row}")
        
        # Ensure the first cell has the correct header for all sheets
        if req_level_row is not None and sheet_df.iloc[req_level_row, 0] != '# requirement_level_code':
            sheet_df.iloc[req_level_row, 0] = '# requirement_level_code'
        
        # Filter by requirement level if req_level_row exists
        if req_level_row is not None:
            req_lev2rm = [level for level in ['M', 'HR', 'R', 'O'] if level not in req_lev]
            cols_to_drop = []
            
            for col_idx in range(sheet_df.shape[1]):
                if sheet_df.iloc[req_level_row, col_idx] in req_lev2rm:
                    cols_to_drop.append(col_idx)
            
            # Drop columns with unwanted requirement levels
            if cols_to_drop:
                sheet_df = sheet_df.drop(columns=sheet_df.columns[cols_to_drop])
        
        # Handle ampData sheet for multiple assay names
        if sheet_name == 'ampData' and len(assay_name) > 1:
            # For ampData with multiple assays, we might need to add assay-specific columns
            # This would depend on the exact requirements, but for now we'll just note it
            pass
        
        # Convert to list of lists for gspread
        data = sheet_df.values.tolist()
        
        # Resize worksheet to accommodate all data (add some buffer)
        rows_needed = len(data) + 20  # Add buffer
        cols_needed = len(data[0]) + 10 if data else 50  # Add buffer
        worksheet.resize(rows=rows_needed, cols=cols_needed)
        
        # Update the worksheet with all data at once
        worksheet.update("A1", data)
        
        # Apply formatting if term_name_row exists
        if term_name_row is not None:
            # Format header row (term_name row)
            header_format = gsf.CellFormat(textFormat=gsf.TextFormat(bold=True))
            gsf.format_cell_range(worksheet, f"{term_name_row+1}:{term_name_row+1}", header_format)
        
        # Apply requirement level colors if req_level_row exists
        if req_level_row is not None:
            for req_code, color_style in color_styles.items():
                # Find cells with this requirement level
                for col_idx in range(len(data[0])):
                    if req_level_row < len(data) and col_idx < len(data[req_level_row]):
                        if data[req_level_row][col_idx] == req_code:
                            # Apply color to this cell
                            gsf.format_cell_range(
                                worksheet, 
                                f"{chr(65 + col_idx)}{req_level_row+1}", 
                                color_style
                            )
        
        # Prepare all comments in a dictionary to add them in a batch
        comments = {}
        if term_name_row is not None:
            print(f"Preparing comments for {sheet_name} sheet...")
            for col_idx in range(len(data[0])):
                if col_idx < len(data[term_name_row]):
                    term_name = data[term_name_row][col_idx]
                    
                    if term_name and isinstance(term_name, str) and not term_name.startswith('#'):
                        # Get requirement level if available
                        req_level = ""
                        if req_level_row is not None and col_idx < len(data[req_level_row]):
                            req_level = data[req_level_row][col_idx]
                        
                        # Get description if available
                        description = ""
                        if description_row is not None and col_idx < len(data[description_row]):
                            description = data[description_row][col_idx]
                        
                        # If no description in the sheet, try to get it from the input dataframe
                        if not description:
                            field_info = input_df[input_df['term_name'] == term_name]
                            if not field_info.empty and 'description' in field_info.columns:
                                description = field_info['description'].iloc[0]
                        
                        # Create comment text
                        comment_text = ""
                        if description:
                            comment_text += f"Description: {description}\n"
                        if req_level:
                            comment_text += f"Requirement level: {req_level}\n"
                        
                        # Check if field has controlled vocabulary
                        field_info = input_df[input_df['term_name'] == term_name]
                        if not field_info.empty:
                            if 'controlled_vocabulary_options' in field_info.columns and not pd.isna(field_info['controlled_vocabulary_options'].iloc[0]):
                                vocab_options = field_info['controlled_vocabulary_options'].iloc[0]
                                comment_text += f"Field type: controlled vocabulary ({vocab_options})\n"
                            elif 'fixed_format' in field_info.columns and not pd.isna(field_info['fixed_format'].iloc[0]):
                                fixed_format = field_info['fixed_format'].iloc[0]
                                comment_text += f"Field type: fixed format ({fixed_format})\n"
                        
                        # Add to comments dictionary if we have content
                        if comment_text:
                            cell_address = f"{chr(65 + col_idx)}{term_name_row+1}"
                            comments[cell_address] = comment_text
        
        # Add all comments at once to avoid rate limits
        if comments:
            print(f"Adding {len(comments)} comments to {sheet_name} sheet...")
            try:
                # First ensure all cells have content
                for cell_address, _ in comments.items():
                    col_letter = cell_address[0]
                    row_num = int(cell_address[1:])
                    col_idx = ord(col_letter) - 65
                    worksheet.update_cell(row_num, col_idx + 1, data[row_num-1][col_idx])
                
                # Wait to avoid rate limits
                time.sleep(2)
                
                # Now add notes in batches to avoid rate limits
                batch_size = 5
                comment_items = list(comments.items())
                
                for i in range(0, len(comment_items), batch_size):
                    batch = comment_items[i:i+batch_size]
                    for cell_address, comment_text in batch:
                        try:
                            worksheet.insert_note(cell_address, comment_text)
                        except Exception as e:
                            print(f"Error adding comment to {cell_address}: {e}")
                    # Wait between batches to avoid rate limits
                    time.sleep(2)
            except Exception as e:
                print(f"Error adding comments to {sheet_name}: {e}")
        
        # Prepare data validation (dropdowns) for fields with controlled vocabulary
        validation_requests = []
        if term_name_row is not None:
            for col_idx in range(len(data[0])):
                if col_idx < len(data[term_name_row]):
                    term_name = data[term_name_row][col_idx]
                    
                    if term_name and isinstance(term_name, str) and not term_name.startswith('#'):
                        # Check if this field has controlled vocabulary
                        field_info = input_df[input_df['term_name'] == term_name]
                        
                        if not field_info.empty and 'controlled_vocabulary_options' in field_info.columns and not pd.isna(field_info['controlled_vocabulary_options'].iloc[0]):
                            vocab_options = field_info['controlled_vocabulary_options'].iloc[0]
                            
                            # Look for this term in vocab_df
                            vocab_row = vocab_df[vocab_df['term_name'] == term_name]
                            
                            if not vocab_row.empty:
                                # Get all non-empty values from columns that start with 'vocab'
                                vocab_cols = [col for col in vocab_df.columns if col.startswith('vocab')]
                                vocab_values = []
                                
                                for col in vocab_cols:
                                    if col in vocab_row.columns and not pd.isna(vocab_row[col].iloc[0]) and vocab_row[col].iloc[0] != '':
                                        value = vocab_row[col].iloc[0]
                                        # Skip values that contain pipe characters (these are likely the full list of options)
                                        if '|' not in value:
                                            vocab_values.append(value)
                                
                                if vocab_values:
                                    print(f"Adding dropdown for {term_name} with values: {vocab_values}")
                                    # Create validation rule
                                    validation_rule = {
                                        "condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": val} for val in vocab_values]},
                                        "showCustomUi": True,
                                        "strict": False
                                    }
                                    
                                    # Define the range for validation (all cells below the header in this column)
                                    grid_range = {
                                        "sheetId": worksheet.id,
                                        "startRowIndex": term_name_row + 1,
                                        "endRowIndex": rows_needed,
                                        "startColumnIndex": col_idx,
                                        "endColumnIndex": col_idx + 1
                                    }
                                    
                                    # Add to validation requests
                                    validation_requests.append({
                                        "setDataValidation": {
                                            "range": grid_range,
                                            "rule": validation_rule
                                        }
                                    })
        
        # Apply all validations in a single batch request
        if validation_requests:
            try:
                # Wait to avoid rate limits
                time.sleep(2)
                body = {"requests": validation_requests}
                worksheet.spreadsheet.batch_update(body)
            except Exception as e:
                print(f"Error applying data validation to {sheet_name}: {e}")
        
        # Wait a moment to avoid rate limits
        time.sleep(2)
