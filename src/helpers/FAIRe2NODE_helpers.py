"""
Module for FAIRe2NODE helper functions.
These functions support the conversion of FAIReSheets templates to NODE format.
"""

import pandas as pd
import gspread
import time
import os
import numpy as np
import gspread_formatting as gsf

def get_bioinformatics_fields(noaa_checklist_path):
    """
    Get list of bioinformatics fields from the NOAA checklist.
    
    Args:
        noaa_checklist_path (str): Path to the NOAA checklist Excel file
        
    Returns:
        list: List of term names that belong to the Bioinformatics section
    """
    try:
        # Read the checklist sheet
        input_df = pd.read_excel(noaa_checklist_path, sheet_name='checklist')
        
        # Get all fields where section is 'Bioinformatics' (lowercase column name)
        bioinfo_fields = input_df[input_df['section'] == 'Bioinformatics']['term_name'].tolist()
        
        return bioinfo_fields
    except Exception as e:
        raise Exception(f"Error reading NOAA checklist: {e}")

def remove_bioinfo_fields_from_project_metadata(worksheet, bioinfo_fields):
    """
    Remove bioinformatics fields from projectMetadata sheet.
    
    Args:
        worksheet (gspread.Worksheet): The projectMetadata worksheet
        bioinfo_fields (list): List of term names to remove
    """
    try:
        # Get all data from the worksheet
        data = worksheet.get_all_values()
        if not data:
            return
            
        # Find the term_name column index
        headers = data[0]
        print(f"Headers in projectMetadata sheet: {headers}")  # Debug
        
        term_name_col = headers.index('term_name')
        
        # Find the project_level column index (this is where dropdowns go)
        project_level_col = None
        for i, header in enumerate(headers):
            if header == 'project_level':
                project_level_col = i
                break
                
        if project_level_col is None:
            print("Warning: Could not find 'project_level' column in projectMetadata sheet")
            return
            
        # Find rows to delete (1-based indexing for worksheet operations)
        rows_to_delete = []
        for i, row in enumerate(data[1:], start=2):  # Start from 2 to skip header
            if row[term_name_col] in bioinfo_fields:
                rows_to_delete.append(i)
        
        # Prepare batch delete request for rows
        batch_requests = []
        for row_idx in sorted(rows_to_delete, reverse=True):
            batch_requests.append({
                "deleteDimension": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "ROWS",
                        "startIndex": row_idx - 1,  # Convert to 0-based
                        "endIndex": row_idx
                    }
                }
            })
        
        # Add request to delete columns after column D (project_level)
        if project_level_col is not None and project_level_col + 1 < len(headers):
            # Delete all columns after project_level column
            batch_requests.append({
                "deleteDimension": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": project_level_col + 1,  # Start after project_level column
                        "endIndex": len(headers)  # Delete all remaining columns
                    }
                }
            })
            print(f"Removing columns {project_level_col + 1} to {len(headers)} from projectMetadata sheet")
        
        # Execute batch delete
        if batch_requests:
            try:
                worksheet.spreadsheet.batch_update({'requests': batch_requests})
            except gspread.exceptions.APIError as e:
                if "429" in str(e):  # Rate limit error
                    print("Warning: Hit API rate limit. Waiting 60 seconds before retrying...")
                    time.sleep(60)
                    worksheet.spreadsheet.batch_update({'requests': batch_requests})
                else:
                    raise
        
        # Now we need to restore the dropdowns
        # First, get the updated data after deletion
        updated_data = worksheet.get_all_values()
        
        # Use the NOAA checklist for vocabulary data
        import os
        import pandas as pd
        
        noaa_checklist_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 
                                         'input', 'FAIRe_NOAA_checklist_v1.0.xlsx')
        
        # Read the checklist sheet
        checklist_df = pd.read_excel(noaa_checklist_path, sheet_name='checklist')
        
        # Prepare batch validation requests
        validation_requests = []
        
        # For each row in the updated sheet
        for i, row in enumerate(updated_data[1:], start=2):  # Skip header row
            term_name = row[term_name_col]
            
            # Find this term in the checklist dataframe
            term_row = checklist_df[checklist_df['term_name'] == term_name]
            if not term_row.empty and 'controlled_vocabulary_options' in term_row.columns:
                vocab_str = term_row.iloc[0]['controlled_vocabulary_options']
                if pd.notna(vocab_str) and vocab_str:
                    # Split the controlled vocabulary string by pipe character
                    values = [v.strip() for v in str(vocab_str).split('|')]
                    
                    if values:
                        # Add data validation for this cell
                        validation_requests.append({
                            "setDataValidation": {
                                "range": {
                                    "sheetId": worksheet.id,
                                    "startRowIndex": i - 1,  # 0-based
                                    "endRowIndex": i,
                                    "startColumnIndex": project_level_col,
                                    "endColumnIndex": project_level_col + 1
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
        
        # Execute batch validation update
        if validation_requests:
            try:
                worksheet.spreadsheet.batch_update({"requests": validation_requests})
            except gspread.exceptions.APIError as e:
                if "429" in str(e):  # Rate limit error
                    print("Warning: Hit API rate limit. Waiting 60 seconds before retrying...")
                    time.sleep(60)
                    worksheet.spreadsheet.batch_update({"requests": validation_requests})
                else:
                    raise
        
        # Remove the last 5 rows from the projectMetadata sheet
        print("Removing the last 5 rows from projectMetadata sheet...")
        # Get the current row count after previous operations
        current_data = worksheet.get_all_values()
        total_rows = len(current_data)
        
        if total_rows > 5:  # Only proceed if there are at least 6 rows (to keep the header)
            # Try direct row deletion instead of batch update
            start_row = total_rows - 4  # 1-based indexing for delete_rows
            worksheet.delete_rows(start_row, total_rows)
            print(f"Successfully removed the last 5 rows from projectMetadata sheet")
        else:
            print(f"Warning: Cannot remove 5 rows as the sheet only has {total_rows} rows")
                    
    except Exception as e:
        raise Exception(f"Error removing bioinformatics fields from projectMetadata: {e}")

def remove_bioinfo_fields_from_experiment_metadata(worksheet, bioinfo_fields):
    """
    Remove bioinformatics fields from experimentRunMetadata sheet.
    
    Args:
        worksheet (gspread.Worksheet): The experimentRunMetadata worksheet
        bioinfo_fields (list): List of term names to remove
    """
    try:
        # Get all data from the worksheet
        data = worksheet.get_all_values()
        if not data:
            return
            
        # Find columns to delete (1-based indexing for worksheet operations)
        cols_to_delete = []
        for i, term in enumerate(data[2]):  # Row 3 (index 2) contains term names
            if term in bioinfo_fields:
                cols_to_delete.append(i + 1)  # Convert to 1-based column index
        
        if not cols_to_delete:
            return
            
        # Prepare batch delete request
        # Note: We need to delete from right to left to maintain correct indices
        batch_requests = []
        for col_idx in sorted(cols_to_delete, reverse=True):
            batch_requests.append({
                "deleteDimension": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": col_idx - 1,  # Convert to 0-based
                        "endIndex": col_idx
                    }
                }
            })
        
        # Execute batch delete
        if batch_requests:
            try:
                worksheet.spreadsheet.batch_update({'requests': batch_requests})
            except gspread.exceptions.APIError as e:
                if "429" in str(e):  # Rate limit error
                    print("Warning: Hit API rate limit. Waiting 60 seconds before retrying...")
                    time.sleep(60)
                    worksheet.spreadsheet.batch_update({'requests': batch_requests})
                else:
                    raise
                    
    except Exception as e:
        raise Exception(f"Error removing bioinformatics fields from experimentRunMetadata: {e}") 
    

# Part 2: Add NOAA fields to the sheets

def get_noaa_fields(noaa_checklist_path, sheet_type):
    """
    Get fields from the NOAA checklist that have the specified NOAA prefix in data_type.
    
    Args:
        noaa_checklist_path (str): Path to the NOAA checklist Excel file
        sheet_type (str): Type of sheet to get fields for (e.g., 'NOAAprojectMetadata')
        
    Returns:
        pandas.DataFrame: DataFrame containing rows with the specified NOAA prefix
    """
    try:
        # Read the checklist sheet
        input_df = pd.read_excel(noaa_checklist_path, sheet_name='checklist')
        
        # Filter rows where data_type contains the specified NOAA prefix
        # This handles cases where multiple values are in the data_type column
        # separated by pipes or other delimiters
        noaa_fields = input_df[input_df['data_type'].apply(
            lambda x: isinstance(x, str) and sheet_type in [t.strip() for t in str(x).split('|')]
        )]
        
        return noaa_fields
    except Exception as e:
        raise Exception(f"Error getting NOAA fields: {e}")


def add_noaa_fields_to_project_metadata(worksheet, noaa_fields):
    """
    Add NOAA fields to projectMetadata sheet.
    
    Args:
        worksheet (gspread.Worksheet): The projectMetadata worksheet
        noaa_fields (pandas.DataFrame): DataFrame containing NOAA fields to add
    """
    try:
        import pandas as pd
        import gspread
        import gspread_formatting as gsf
        import numpy as np
        import time
        
        # Replace NaN values with empty strings
        noaa_fields = noaa_fields.fillna('')
        
        # Get all data from the worksheet
        data = worksheet.get_all_values()
        if not data:
            return
            
        # Find the term_name column index
        headers = data[0]
        
        # Convert data to DataFrame for easier manipulation
        sheet_df = pd.DataFrame(data[1:], columns=headers)
        
        # Prepare new rows to add
        new_rows = []
        for _, row in noaa_fields.iterrows():
            new_row = {}
            for col in headers:
                new_row[col] = ''
            
            # Fill in the values we have
            new_row['term_name'] = row['term_name']
            new_row['requirement_level_code'] = row['requirement_level_code']
            new_row['section'] = row['section']
            new_row['description'] = row['description'] if 'description' in row and row['description'] != '' else ''
            
            # Add controlled vocabulary if available
            if 'controlled_vocabulary' in row and row['controlled_vocabulary'] != '':
                new_row['controlled_vocabulary'] = row['controlled_vocabulary']
            
            new_rows.append(new_row)
        
        # Convert to DataFrame and append to existing data
        new_rows_df = pd.DataFrame(new_rows)
        updated_df = pd.concat([sheet_df, new_rows_df], ignore_index=True)
        
        # Replace any NaN values with empty strings
        updated_df = updated_df.fillna('')
        
        # Convert back to list of lists for updating the sheet
        updated_data = [headers] + updated_df.values.tolist()
        
        # Update the worksheet with all data at once
        worksheet.resize(rows=len(updated_data) + 10, cols=len(headers) + 5)  # Add buffer
        worksheet.update("A1", updated_data)
        
        # Define color styles for requirement levels
        color_styles = {
            "M": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#E26B0A")),  # Mandatory - Orange
            "HR": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#FFCC00")), # Highly recommended - Yellow
            "R": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#FFFF99")),  # Recommended - Light yellow
            "O": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#CCFF99"))   # Optional - Light green
        }
        
        # Format cells based on requirement level
        req_level_col = headers.index('requirement_level_code')
        term_name_col = headers.index('term_name')
        
        # Batch requests for formatting
        batch_requests = []
        
        # Apply formatting to new rows
        for i, row in enumerate(updated_data[1:], start=1):
            # Skip if no requirement level
            if req_level_col >= len(row) or not row[req_level_col]:
                continue
                
            req_level = row[req_level_col]
            if req_level in color_styles:
                # Add color formatting for requirement level
                batch_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": i,
                            "endRowIndex": i + 1,
                            "startColumnIndex": req_level_col,
                            "endColumnIndex": req_level_col + 1
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": color_styles[req_level].backgroundColor.red,
                                    "green": color_styles[req_level].backgroundColor.green,
                                    "blue": color_styles[req_level].backgroundColor.blue
                                }
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })
                
            # Bold the term name
            batch_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": i,
                        "endRowIndex": i + 1,
                        "startColumnIndex": term_name_col,
                        "endColumnIndex": term_name_col + 1
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
        
        # Apply batch formatting
        if batch_requests:
            worksheet.spreadsheet.batch_update({"requests": batch_requests})
            
        # Add descriptions as notes
        note_requests = []
        desc_col = headers.index('description') if 'description' in headers else None
        
        if desc_col is not None:
            for i, row in enumerate(updated_data[1:], start=1):
                if desc_col < len(row) and row[desc_col]:
                    note_requests.append({
                        "updateCells": {
                            "range": {
                                "sheetId": worksheet.id,
                                "startRowIndex": i,
                                "endRowIndex": i + 1,
                                "startColumnIndex": term_name_col,
                                "endColumnIndex": term_name_col + 1
                            },
                            "rows": [{
                                "values": [{
                                    "note": row[desc_col]
                                }]
                            }],
                            "fields": "note"
                        }
                    })
        
        # Apply notes
        if note_requests:
            worksheet.spreadsheet.batch_update({"requests": note_requests})
            
        # Add data validation for controlled vocabulary
        validation_requests = []
        cv_col = headers.index('controlled_vocabulary') if 'controlled_vocabulary' in headers else None
        
        if cv_col is not None:
            for i, row in enumerate(updated_data[1:], start=1):
                if cv_col < len(row) and row[cv_col]:
                    # Parse the controlled vocabulary values
                    cv_values = [v.strip() for v in row[cv_col].split('|') if v.strip()]
                    if cv_values:
                        validation_requests.append({
                            "setDataValidation": {
                                "range": {
                                    "sheetId": worksheet.id,
                                    "startRowIndex": i,
                                    "endRowIndex": i + 1,
                                    "startColumnIndex": cv_col,
                                    "endColumnIndex": cv_col + 1
                                },
                                "rule": {
                                    "condition": {
                                        "type": "ONE_OF_LIST",
                                        "values": [{"userEnteredValue": v} for v in cv_values]
                                    },
                                    "showCustomUi": True,
                                    "strict": False
                                }
                            }
                        })
        
        # Apply data validation
        if validation_requests:
            worksheet.spreadsheet.batch_update({"requests": validation_requests})
        
        print(f"Added {len(new_rows)} NOAA fields to projectMetadata sheet with formatting")
        
    except Exception as e:
        raise Exception(f"Error adding NOAA fields to projectMetadata: {e}")


def add_noaa_fields_to_experiment_metadata(worksheet, noaa_fields):
    """
    Add NOAA fields to experimentRunMetadata sheet.
    
    Args:
        worksheet (gspread.Worksheet): The experimentRunMetadata worksheet
        noaa_fields (pandas.DataFrame): DataFrame containing NOAA fields to add
    """
    try:
        import pandas as pd
        import numpy as np
        import gspread_formatting as gsf
        
        # Replace NaN values with empty strings
        noaa_fields = noaa_fields.fillna('')
        
        # Get all data from the worksheet
        data = worksheet.get_all_values()
        if not data:
            return
            
        # Find key rows
        term_name_row = None
        req_level_row = None
        section_row = None
        description_row = None
        
        for idx, row in enumerate(data):
            if '# requirement_level_code' in row:
                req_level_row = idx
            if '# section' in row:
                section_row = idx
            if '# description' in row:
                description_row = idx
            # The term name row is typically the last row with actual field names
            if any(col for col in row if isinstance(col, str) and not col.startswith('#')):
                term_name_row = idx
        
        if term_name_row is None:
            # Default to row 2 if not found
            term_name_row = 2
            
        # Convert data to DataFrame for easier manipulation
        sheet_df = pd.DataFrame(data)
        
        # Prepare new columns to add
        new_cols = []
        for _, row in noaa_fields.iterrows():
            # Create a new column for the NOAA field
            new_col_idx = sheet_df.shape[1]
            sheet_df[new_col_idx] = ''
            
            # Set term name
            sheet_df.iloc[term_name_row, new_col_idx] = row['term_name']
            
            # Set requirement level and section if those rows exist
            if req_level_row is not None:
                sheet_df.iloc[req_level_row, new_col_idx] = row['requirement_level_code']
            if section_row is not None:
                sheet_df.iloc[section_row, new_col_idx] = row['section']
            if description_row is not None and 'description' in row and row['description'] != '':
                sheet_df.iloc[description_row, new_col_idx] = row['description']
            
            new_cols.append(new_col_idx)
        
        # Replace any NaN values with empty strings
        sheet_df = sheet_df.fillna('')
        
        # Convert back to list of lists for updating the sheet
        updated_data = sheet_df.values.tolist()
        
        # Update the worksheet with all data at once
        worksheet.resize(rows=len(updated_data) + 10, cols=len(updated_data[0]) + 5)  # Add buffer
        worksheet.update("A1", updated_data)
        
        # Define color styles for requirement levels
        color_styles = {
            "M": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#E26B0A")),  # Mandatory - Orange
            "HR": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#FFCC00")), # Highly recommended - Yellow
            "R": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#FFFF99")),  # Recommended - Light yellow
            "O": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#CCFF99"))   # Optional - Light green
        }
        
        # Format cells based on requirement level
        batch_requests = []
        
        # Apply color formatting to requirement level cells
        if req_level_row is not None:
            for col_idx in new_cols:
                req_level = sheet_df.iloc[req_level_row, col_idx]
                if req_level in color_styles:
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
                                        "red": color_styles[req_level].backgroundColor.red,
                                        "green": color_styles[req_level].backgroundColor.green,
                                        "blue": color_styles[req_level].backgroundColor.blue
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.backgroundColor"
                        }
                    })
                    
        # Bold the term names
        for col_idx in new_cols:
            batch_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": term_name_row,
                        "endRowIndex": term_name_row + 1,
                        "startColumnIndex": col_idx,
                        "endColumnIndex": col_idx + 1
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
        
        # Apply batch formatting
        if batch_requests:
            worksheet.spreadsheet.batch_update({"requests": batch_requests})
            
        # Add notes to term names
        note_requests = []
        for col_idx in new_cols:
            term_name = sheet_df.iloc[term_name_row, col_idx]
            term_rows = noaa_fields[noaa_fields['term_name'] == term_name]
            if not term_rows.empty:
                term_info = term_rows.iloc[0]
                if 'description' in term_info and term_info['description']:
                    note_requests.append({
                        "updateCells": {
                            "range": {
                                "sheetId": worksheet.id,
                                "startRowIndex": term_name_row,
                                "endRowIndex": term_name_row + 1,
                                "startColumnIndex": col_idx,
                                "endColumnIndex": col_idx + 1
                            },
                            "rows": [{
                                "values": [{
                                    "note": term_info['description']
                                }]
                            }],
                            "fields": "note"
                        }
                    })
        
        # Apply notes
        if note_requests:
            worksheet.spreadsheet.batch_update({"requests": note_requests})
            
        # Add data validation for controlled vocabulary
        validation_requests = []
        for col_idx in new_cols:
            term_name = sheet_df.iloc[term_name_row, col_idx]
            term_rows = noaa_fields[noaa_fields['term_name'] == term_name]
            if not term_rows.empty:
                term_info = term_rows.iloc[0]
                if 'controlled_vocabulary' in term_info and term_info['controlled_vocabulary']:
                    # Parse the controlled vocabulary values
                    cv_values = [v.strip() for v in str(term_info['controlled_vocabulary']).split('|') if v.strip()]
                    if cv_values:
                        validation_requests.append({
                            "setDataValidation": {
                                "range": {
                                    "sheetId": worksheet.id,
                                    "startRowIndex": term_name_row + 1,  # Start from the row after term names
                                    "endRowIndex": len(updated_data),
                                    "startColumnIndex": col_idx,
                                    "endColumnIndex": col_idx + 1
                                },
                                "rule": {
                                    "condition": {
                                        "type": "ONE_OF_LIST",
                                        "values": [{"userEnteredValue": v} for v in cv_values]
                                    },
                                    "showCustomUi": True,
                                    "strict": False
                                }
                            }
                        })
        
        # Apply data validation
        if validation_requests:
            worksheet.spreadsheet.batch_update({"requests": validation_requests})
        
        print(f"Added {len(new_cols)} NOAA fields to experimentRunMetadata sheet with formatting")
        
    except Exception as e:
        raise Exception(f"Error adding NOAA fields to experimentRunMetadata: {e}")


def add_noaa_fields_to_sample_metadata(worksheet, noaa_fields):
    """
    Add NOAA fields to sampleMetadata sheet.
    
    Args:
        worksheet (gspread.Worksheet): The sampleMetadata worksheet
        noaa_fields (pandas.DataFrame): DataFrame containing NOAA fields to add
    """
    try:
        import pandas as pd
        import numpy as np
        import gspread_formatting as gsf
        
        # Replace NaN values with empty strings
        noaa_fields = noaa_fields.fillna('')
        
        # Get all data from the worksheet
        data = worksheet.get_all_values()
        if not data:
            return
            
        # Find key rows
        term_name_row = None
        req_level_row = None
        section_row = None
        description_row = None
        
        for idx, row in enumerate(data):
            if '# requirement_level_code' in row:
                req_level_row = idx
            if '# section' in row:
                section_row = idx
            if '# description' in row:
                description_row = idx
            # The term name row is typically the last row with actual field names
            if any(col for col in row if isinstance(col, str) and not col.startswith('#')):
                term_name_row = idx
        
        if term_name_row is None or req_level_row is None or section_row is None:
            raise Exception("Could not find term name, requirement level, or section row in sampleMetadata")
        
        term_name_row = term_name_row - 1  # Move term names up by one row
        
        # Convert data to DataFrame for easier manipulation
        sheet_df = pd.DataFrame(data)
        
        # Prepare new columns to add
        new_cols = []
        for _, row in noaa_fields.iterrows():
            # Create a new column for the NOAA field
            new_col_idx = sheet_df.shape[1]
            sheet_df[new_col_idx] = ''
            
            # Set term name, requirement level, and section
            sheet_df.iloc[term_name_row, new_col_idx] = row['term_name']
            sheet_df.iloc[req_level_row, new_col_idx] = row['requirement_level_code']
            sheet_df.iloc[section_row, new_col_idx] = row['section']
            
            # Set description if available
            if description_row is not None and 'description' in row and row['description'] != '':
                sheet_df.iloc[description_row, new_col_idx] = row['description']
            
            new_cols.append(new_col_idx)
        
        # Replace any NaN values with empty strings
        sheet_df = sheet_df.fillna('')
        
        # Convert back to list of lists for updating the sheet
        updated_data = sheet_df.values.tolist()
        
        # Update the worksheet with all data at once
        worksheet.resize(rows=len(updated_data) + 10, cols=len(updated_data[0]) + 5)  # Add buffer
        worksheet.update("A1", updated_data)
        
        # Define color styles for requirement levels
        color_styles = {
            "M": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#E26B0A")),  # Mandatory - Orange
            "HR": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#FFCC00")), # Highly recommended - Yellow
            "R": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#FFFF99")),  # Recommended - Light yellow
            "O": gsf.CellFormat(backgroundColor=gsf.Color.fromHex("#CCFF99"))   # Optional - Light green
        }
        
        # Format cells based on requirement level
        batch_requests = []
        
        # Apply color formatting to requirement level cells
        for col_idx in new_cols:
            req_level = sheet_df.iloc[req_level_row, col_idx]
            if req_level in color_styles:
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
                                    "red": color_styles[req_level].backgroundColor.red,
                                    "green": color_styles[req_level].backgroundColor.green,
                                    "blue": color_styles[req_level].backgroundColor.blue
                                }
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })
                
            # Bold the term name
            batch_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": term_name_row,
                        "endRowIndex": term_name_row + 1,
                        "startColumnIndex": col_idx,
                        "endColumnIndex": col_idx + 1
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
        
        # Apply batch formatting
        if batch_requests:
            worksheet.spreadsheet.batch_update({"requests": batch_requests})
            
        # Add notes to term names
        note_requests = []
        for col_idx in new_cols:
            term_name = sheet_df.iloc[term_name_row, col_idx]
            term_rows = noaa_fields[noaa_fields['term_name'] == term_name]
            if not term_rows.empty:
                term_info = term_rows.iloc[0]
                if 'description' in term_info and term_info['description']:
                    note_requests.append({
                        "updateCells": {
                            "range": {
                                "sheetId": worksheet.id,
                                "startRowIndex": term_name_row,
                                "endRowIndex": term_name_row + 1,
                                "startColumnIndex": col_idx,
                                "endColumnIndex": col_idx + 1
                            },
                            "rows": [{
                                "values": [{
                                    "note": term_info['description']
                                }]
                            }],
                            "fields": "note"
                        }
                    })
        
        # Apply notes
        if note_requests:
            worksheet.spreadsheet.batch_update({"requests": note_requests})
            
        # Add data validation for controlled vocabulary
        validation_requests = []
        for col_idx in new_cols:
            term_name = sheet_df.iloc[term_name_row, col_idx]
            term_rows = noaa_fields[noaa_fields['term_name'] == term_name]
            if not term_rows.empty:
                term_info = term_rows.iloc[0]
                if 'controlled_vocabulary' in term_info and term_info['controlled_vocabulary']:
                    # Parse the controlled vocabulary values
                    cv_values = [v.strip() for v in str(term_info['controlled_vocabulary']).split('|') if v.strip()]
                    if cv_values:
                        validation_requests.append({
                            "setDataValidation": {
                                "range": {
                                    "sheetId": worksheet.id,
                                    "startRowIndex": term_name_row + 1,  # Start from the row after term names
                                    "endRowIndex": len(updated_data),
                                    "startColumnIndex": col_idx,
                                    "endColumnIndex": col_idx + 1
                                },
                                "rule": {
                                    "condition": {
                                        "type": "ONE_OF_LIST",
                                        "values": [{"userEnteredValue": v} for v in cv_values]
                                    },
                                    "showCustomUi": True,
                                    "strict": False
                                }
                            }
                        })
        
        # Apply data validation
        if validation_requests:
            worksheet.spreadsheet.batch_update({"requests": validation_requests})
        
        print(f"Added {len(new_cols)} NOAA fields to sampleMetadata sheet with formatting")
        
    except Exception as e:
        raise Exception(f"Error adding NOAA fields to sampleMetadata: {e}")

def remove_taxa_sheets(spreadsheet):
    """
    Remove taxaRaw and taxaFinal sheets from the spreadsheet.
    
    Args:
        spreadsheet (gspread.Spreadsheet): The Google Spreadsheet object
    """
    try:
        # Get the worksheets
        taxa_raw = spreadsheet.worksheet("taxaRaw")
        taxa_final = spreadsheet.worksheet("taxaFinal")
        
        # Delete the worksheets
        spreadsheet.del_worksheet(taxa_raw)
        spreadsheet.del_worksheet(taxa_final)
        
        print("Successfully removed taxaRaw and taxaFinal sheets")
    except Exception as e:
        raise Exception(f"Error removing taxa sheets: {e}")

def create_analysis_metadata_sheets(spreadsheet, config):
    """
    Create analysisMetadata Google Sheets for each analysis run name in the config.
    
    This function creates a new Google Sheet named 'analysisMetadata_<analysis_run_name>'
    for each analysis run name specified in the NOAA_config.yaml file. If no analysis
    run names are provided, it creates a single generic 'analysisMetadata' sheet.
    
    Args:
        spreadsheet (gspread.Spreadsheet): The Google Spreadsheet object
        config (dict): Configuration loaded from NOAA_config.yaml
        
    Returns:
        dict: Dictionary mapping analysis run names to their worksheet objects
        
    Raises:
        Exception: If there's an error creating the sheets
    """
    try:
        # Get project_id from config
        project_id = config.get('project_id')
        if not project_id:
            raise ValueError("project_id not found in NOAA_config.yaml")
        
        # Get analysis run names from config
        analysis_runs = config.get('analysis_run_name', {})
        
        # Dictionary to store created worksheets
        analysis_worksheets = {}
        
        # If no analysis runs are specified, create a single generic analysisMetadata sheet
        if not analysis_runs:
            print("No analysis run names specified in config. Creating a single analysisMetadata sheet...")
            sheet_name = "analysisMetadata_<analysis_run_name>"
            try:
                worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=100)
                analysis_worksheets[sheet_name] = worksheet
                print(f"Created {sheet_name} sheet")
            except gspread.exceptions.APIError as e:
                if "429" in str(e):  # Rate limit error
                    print("Warning: Hit API rate limit. Waiting 60 seconds before retrying...")
                    time.sleep(60)
                    worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=100)
                    analysis_worksheets[sheet_name] = worksheet
                    print(f"Created {sheet_name} sheet")
                else:
                    raise
        else:
            # Create a sheet for each analysis run name
            for analysis_run_name in analysis_runs:
                sheet_name = f"analysisMetadata_{analysis_run_name}"
                try:
                    worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=100)
                    analysis_worksheets[analysis_run_name] = worksheet
                    print(f"Created {sheet_name} sheet")
                except gspread.exceptions.APIError as e:
                    if "429" in str(e):  # Rate limit error
                        print("Warning: Hit API rate limit. Waiting 60 seconds before retrying...")
                        time.sleep(60)
                        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=100)
                        analysis_worksheets[analysis_run_name] = worksheet
                        print(f"Created {sheet_name} sheet")
                    else:
                        raise
        
        return analysis_worksheets
    
    except Exception as e:
        raise Exception(f"Error creating analysis metadata sheets: {e}")