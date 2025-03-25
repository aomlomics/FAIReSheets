"""
Module for creating targeted assay sheets in FAIReSheets.

This module contains functions to create sheets for targeted assay data including
standard curves, quantitative data, and amplification data.
"""

import pandas as pd
import gspread_formatting as gsf
import time

def create_targeted_sheets(worksheets, sheet_names, full_temp_file_name, input_df, req_lev, 
                          color_styles, vocab_df, project_id, assay_name):
    """Create and format targeted assay sheets."""
    
    # Process each sheet
    for sheet_name in sheet_names:
        try:
            # Get the data from the template
            sheet_df = pd.read_excel(full_temp_file_name, sheet_name=sheet_name)
            
            # Replace NaN values with empty strings immediately after loading
            sheet_df = sheet_df.fillna('')
            
            # Find the column positions
            # Use loc to avoid SettingWithCopyWarning, then get the index
            try:
                term_name_row = sheet_df.loc[sheet_df.iloc[:, 0] == 'term_name'].index[0]
                req_level_row = sheet_df.loc[sheet_df.iloc[:, 0] == 'requirement_level_code'].index[0]
                section_row = sheet_df.loc[sheet_df.iloc[:, 0] == 'section'].index[0]
                description_row = sheet_df.loc[sheet_df.iloc[:, 0] == 'description'].index[0]
            except (IndexError, KeyError) as e:
                # This is a critical error, so we'll keep this message
                print(f"Error finding template rows: {e}")
                continue
            
            # Extract the headers and create a new DataFrame for the sheet content
            headers = sheet_df.iloc[term_name_row].values
            req_levels = sheet_df.iloc[req_level_row].values
            sections = sheet_df.iloc[section_row].values
            descriptions = sheet_df.iloc[description_row].values
            
            # Find the first data row
            data_start_row = description_row + 1
            
            # Create a new DataFrame for the output
            output_df = pd.DataFrame(columns=headers)
            
            # Add requirement level, section, and description rows
            output_df.loc['requirement_level_code'] = req_levels
            output_df.loc['section'] = sections
            output_df.loc['description'] = descriptions
            
            # Filter rows based on requirement levels
            if req_lev and 'requirement_level_code' in output_df.index:
                for col in output_df.columns[1:]:  # Skip the first column
                    req_level = output_df.at['requirement_level_code', col]
                    if req_level and req_level not in req_lev:
                        # Remove column by setting values to empty
                        for row in output_df.index:
                            output_df.at[row, col] = ''
            
            # Add project_id and assay name for first assay (if available)
            if 'projectID' in output_df.columns:
                output_df.at['description', 'projectID'] = project_id
            
            if 'assayName' in output_df.columns and assay_name and len(assay_name) > 0:
                output_df.at['description', 'assayName'] = assay_name[0]  # Use first assay name
            
            # Convert DataFrame to list of lists for gspread
            data = [output_df.columns.tolist()]
            for idx in output_df.index:
                data.append(output_df.loc[idx].tolist())
            
            # Resize worksheet
            try:
                rows_needed = len(data) + 10  # Add buffer
                cols_needed = len(data[0]) + 5  # Add buffer
                worksheet = worksheets[sheet_name]
                worksheet.resize(rows=rows_needed, cols=cols_needed)
            except Exception as e:
                # Keep this error as it might explain why a sheet is missing
                print(f"Error resizing worksheet: {e}. Continuing with current size.")
            
            # Update the worksheet
            try:
                worksheet = worksheets[sheet_name]
                worksheet.update("A1", data)
            except Exception as e:
                # Keep this error as it's critical
                print(f"Error updating worksheet: {e}. Skipping this sheet.")
                continue
            
            # Format header row with bold text
            try:
                header_format = gsf.CellFormat(textFormat=gsf.TextFormat(bold=True))
                gsf.format_cell_ranges(worksheet, [('1:1', header_format)])
            except Exception as e:
                # This is non-critical, so we'll skip the error message
                pass
            
            # Format requirement level column with colors based on values
            try:
                # Get the requirement level column
                req_level_col = 2  # Usually column B
                
                # Create a list of individual cell formatting requests
                format_requests = []
                
                for row_idx in range(1, len(data)):
                    # Skip the header row (row 0)
                    if row_idx < len(data) and 0 < req_level_col-1 < len(data[row_idx]):
                        req_level = data[row_idx][req_level_col-1]
                        if req_level in color_styles:
                            cell = f"{chr(64 + req_level_col)}{row_idx+1}"
                            format_requests.append((cell, color_styles[req_level]))
                
                # Apply all formatting at once
                if format_requests:
                    gsf.format_cell_ranges(worksheet, format_requests)
            except Exception as e:
                # This is non-critical, so we'll skip the error message
                pass
            
            # Prepare to add comments
            # Get the description row index (usually row 3)
            desc_row_idx = 3
            
            # Create comments for columns
            comments = []
            for col_idx in range(1, len(headers)):
                if col_idx < len(descriptions):
                    desc = descriptions[col_idx]
                    if desc and desc != '':
                        cell_address = f"{chr(64 + col_idx + 1)}{desc_row_idx}"
                        comments.append((cell_address, desc))
            
            # Apply comments in batches to avoid rate limits
            batch_size = 10
            
            for i in range(0, len(comments), batch_size):
                batch = comments[i:i+batch_size]
                
                # Apply each comment in the batch
                for cell_address, comment_text in batch:
                    try:
                        worksheet.update_note(cell_address, comment_text)
                    except Exception as e:
                        # Skip commenting errors - they don't affect functionality
                        pass
                
                # Avoid Google API rate limits
                if i + batch_size < len(comments):
                    time.sleep(2)
            
            # Add dropdown validations if applicable
            validation_requests = []
            
            # Find columns with vocabulary entries
            for col_idx, header in enumerate(headers):
                if col_idx == 0:  # Skip row header column
                    continue
                    
                # Find vocabulary for this term
                term_name = header
                vocab_row = vocab_df[vocab_df['term_name'] == term_name]
                
                if not vocab_row.empty:
                    # Get the number of options
                    n_options = int(vocab_row.iloc[0]['n_options'])
                    
                    # Get the vocabulary values
                    vocab_values = [str(vocab_row.iloc[0][f'vocab{i+1}']) for i in range(n_options) 
                                   if i+1 <= len(vocab_row.iloc[0]) and f'vocab{i+1}' in vocab_row.iloc[0].index and pd.notna(vocab_row.iloc[0][f'vocab{i+1}'])]
                    
                    if vocab_values:
                        # Create validation for all rows starting from after the description row
                        data_start = desc_row_idx + 1  # 1-indexed
                        
                        # Create validation rule
                        validation_rule = {
                            "setDataValidation": {
                                "range": {
                                    "sheetId": worksheet.id,
                                    "startRowIndex": data_start - 1,  # Convert to 0-indexed
                                    "endRowIndex": rows_needed,
                                    "startColumnIndex": col_idx,
                                    "endColumnIndex": col_idx + 1
                                },
                                "rule": {
                                    "condition": {
                                        "type": "ONE_OF_LIST",
                                        "values": [{"userEnteredValue": v} for v in vocab_values]
                                    },
                                    "showCustomUi": True,
                                    "strict": False
                                }
                            }
                        }
                        validation_requests.append(validation_rule)
            
            # Apply all validations in one batch request
            if validation_requests:
                try:
                    worksheet.spreadsheet.batch_update({"requests": validation_requests})
                except Exception as e:
                    # This is non-critical, so we'll skip the error message
                    # Wait for the rate limit to reset if needed
                    time.sleep(30)
                    
        except Exception as e:
            # This is a critical error, so we'll keep this message
            print(f"Error processing {sheet_name} sheet: {e}")
            continue
