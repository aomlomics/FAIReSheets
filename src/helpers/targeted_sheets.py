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
    
    # Process one sheet at a time with a delay between sheets
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
        
        # Convert to list of lists for gspread
        data = sheet_df.values.tolist()
        
        # Resize worksheet to accommodate all data (add some buffer)
        rows_needed = len(data) + 20  # Add buffer
        cols_needed = len(data[0]) + 10 if data else 50  # Add buffer
        
        try:
            worksheet.resize(rows=rows_needed, cols=cols_needed)
            time.sleep(2)  # Short delay after resize
        except Exception as e:
            print(f"Error resizing worksheet: {e}. Continuing with current size.")
        
        # Update the worksheet with all data at once
        try:
            worksheet.update("A1", data)
            time.sleep(2)  # Short delay after update
        except Exception as e:
            print(f"Error updating worksheet: {e}. Skipping this sheet.")
            continue
        
        # Apply formatting if term_name_row exists
        if term_name_row is not None:
            try:
                # Format header row (term_name row)
                header_format = gsf.CellFormat(textFormat=gsf.TextFormat(bold=True))
                gsf.format_cell_range(worksheet, f"{term_name_row+1}:{term_name_row+1}", header_format)
            except Exception as e:
                print(f"Error formatting header: {e}")
        
        # Apply requirement level colors if req_level_row exists
        if req_level_row is not None:
            for req_code, color_style in color_styles.items():
                try:
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
                except Exception as e:
                    print(f"Error applying colors: {e}")
        
        # Wait after basic formatting to avoid rate limits
        time.sleep(5)
        
        # Collect all comments in a dictionary
        comments = {}
        if term_name_row is not None:
            print(f"Preparing comments for {sheet_name} sheet...")
            for col_idx in range(len(data[0])):
                if col_idx < len(data[term_name_row]):
                    term_name = data[term_name_row][col_idx]
                    
                    if term_name and isinstance(term_name, str) and not term_name.startswith('#'):
                        # Get field information from input_df
                        field_info = input_df[input_df['term_name'] == term_name]
                        
                        # Build comment text
                        comment_text = ""
                        
                        if not field_info.empty:
                            # Add requirement level
                            if 'requirement_level' in field_info.columns and not pd.isna(field_info['requirement_level'].iloc[0]):
                                comment_text += f"Requirement level: {field_info['requirement_level'].iloc[0]}"
                                
                                # Add condition if it exists
                                if 'requirement_level_condition' in field_info.columns and not pd.isna(field_info['requirement_level_condition'].iloc[0]):
                                    comment_text += f" ({field_info['requirement_level_condition'].iloc[0]})"
                                
                                comment_text += "\n"
                            
                            # Add description
                            if 'description' in field_info.columns and not pd.isna(field_info['description'].iloc[0]):
                                comment_text += f"Description: {field_info['description'].iloc[0]}\n"
                            
                            # Add example
                            if 'example' in field_info.columns and not pd.isna(field_info['example'].iloc[0]):
                                comment_text += f"Example: {field_info['example'].iloc[0]}\n"
                            
                            # Add field type
                            if 'term_type' in field_info.columns and not pd.isna(field_info['term_type'].iloc[0]):
                                comment_text += f"Field type: {field_info['term_type'].iloc[0]}"
                                
                                # Add controlled vocabulary options if applicable
                                if 'controlled_vocabulary_options' in field_info.columns and not pd.isna(field_info['controlled_vocabulary_options'].iloc[0]):
                                    vocab_options = field_info['controlled_vocabulary_options'].iloc[0]
                                    comment_text += f" ({vocab_options})\n"
                                elif 'fixed_format' in field_info.columns and not pd.isna(field_info['fixed_format'].iloc[0]):
                                    fixed_format = field_info['fixed_format'].iloc[0]
                                    comment_text += f" ({fixed_format})\n"
                        
                        # Add to comments dictionary if we have content
                        if comment_text:
                            cell_address = f"{chr(65 + col_idx)}{term_name_row+1}"
                            comments[cell_address] = comment_text
        
        # Add comments in batches with delays
        if comments:
            print(f"Adding {len(comments)} comments to {sheet_name} sheet...")
            comment_items = list(comments.items())
            batch_size = 5  # Process 5 comments at a time
            
            for i in range(0, len(comment_items), batch_size):
                # Wait between batches
                if i > 0:
                    time.sleep(5)
                    
                batch = comment_items[i:i+batch_size]
                for cell_address, comment_text in batch:
                    try:
                        worksheet.insert_note(cell_address, comment_text)
                    except Exception as e:
                        print(f"Error adding comment to {cell_address}: {e}")
                        # If we hit a rate limit, wait longer before continuing
                        if "429" in str(e):
                            print("Rate limit hit, waiting 30 seconds...")
                            time.sleep(30)
        
        # Wait after adding comments to avoid rate limits
        time.sleep(5)
        
        # Collect all validation requests
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
        
        # Apply validations in batches with delays
        if validation_requests:
            batch_size = 5  # Process 5 validations at a time
            
            for i in range(0, len(validation_requests), batch_size):
                # Wait between batches
                if i > 0:
                    time.sleep(5)
                    
                batch = validation_requests[i:i+batch_size]
                try:
                    body = {"requests": batch}
                    worksheet.spreadsheet.batch_update(body)
                except Exception as e:
                    print(f"Error applying validation batch: {e}")
                    # If we hit a rate limit, wait longer before continuing
                    if "429" in str(e):
                        print("Rate limit hit, waiting 30 seconds...")
                        time.sleep(30)
        
        # Wait before processing the next sheet
        time.sleep(10)
