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
import webbrowser

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
        
        term_name_col = headers.index('term_name')
        
        # Find the project_level column index (this is where dropdowns go)
        project_level_col = None
        for i, header in enumerate(headers):
            if header == 'project_level':
                project_level_col = i
                break
                
        if project_level_col is None:
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
        
        # Execute batch delete
        if batch_requests:
            try:
                worksheet.spreadsheet.batch_update({'requests': batch_requests})
            except gspread.exceptions.APIError as e:
                if "429" in str(e):  # Rate limit error
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
                                         'input', 'FAIRe_NOAA_checklist_v1.0.2.xlsx')
        
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
                    time.sleep(60)
                    worksheet.spreadsheet.batch_update({"requests": validation_requests})
                else:
                    raise
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
            
            # Add to new rows
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
        project_level_col = headers.index('project_level')
        
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
            
        # Add descriptions as notes and controlled vocabulary dropdowns
        note_requests = []
        validation_requests = []
        
        for i, row in enumerate(updated_data[1:], start=1):
            term_name = row[term_name_col]
            term_info = noaa_fields[noaa_fields['term_name'] == term_name]
            
            if not term_info.empty:
                # Add description as note
                description = term_info.iloc[0]['description'] if 'description' in term_info.columns else ''
                if description:
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
                                    "note": description
                                }]
                            }],
                            "fields": "note"
                        }
                    })
                
                # Add controlled vocabulary dropdown
                cv_options = term_info.iloc[0]['controlled_vocabulary_options'] if 'controlled_vocabulary_options' in term_info.columns else ''
                if pd.notna(cv_options) and cv_options:
                    values = [v.strip() for v in str(cv_options).split('|') if v.strip()]
                    if values:
                        validation_requests.append({
                            "setDataValidation": {
                                "range": {
                                    "sheetId": worksheet.id,
                                    "startRowIndex": i,
                                    "endRowIndex": i + 1,
                                    "startColumnIndex": project_level_col,
                                    "endColumnIndex": project_level_col + 1
                                },
                                "rule": {
                                    "condition": {
                                        "type": "ONE_OF_LIST",
                                        "values": [{"userEnteredValue": v} for v in values]
                                    },
                                    "showCustomUi": True,
                                    "strict": False
                                }
                            }
                        })
        
        # Apply notes
        if note_requests:
            worksheet.spreadsheet.batch_update({"requests": note_requests})
        
        # Apply data validation
        if validation_requests:
            worksheet.spreadsheet.batch_update({"requests": validation_requests})
        
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
            
        # Add notes to term names and controlled vocabulary dropdowns
        note_requests = []
        validation_requests = []
        
        for col_idx in new_cols:
            term_name = sheet_df.iloc[term_name_row, col_idx]
            term_rows = noaa_fields[noaa_fields['term_name'] == term_name]
            if not term_rows.empty:
                term_info = term_rows.iloc[0]
                
                # Add description as note
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
                
                # Add controlled vocabulary dropdown - FIXED VERSION
                if 'controlled_vocabulary_options' in term_info and term_info['controlled_vocabulary_options']:
                    # Parse the controlled vocabulary values
                    cv_values = [v.strip() for v in str(term_info['controlled_vocabulary_options']).split('|') if v.strip()]
                    if cv_values:
                        # Remove the debug print that was interrupting the progress bar
                        # print(f"Adding dropdown for {term_name} with values: {cv_values}")
                        
                        # Apply to all data rows
                        validation_requests.append({
                            "setDataValidation": {
                                "range": {
                                    "sheetId": worksheet.id,
                                    "startRowIndex": term_name_row + 1,  # Start from the row after term names
                                    "endRowIndex": max(term_name_row + 20, len(updated_data)),  # Ensure we have enough rows
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
        
        # Apply notes
        if note_requests:
            worksheet.spreadsheet.batch_update({"requests": note_requests})
        
        # Apply data validation
        if validation_requests:
            worksheet.spreadsheet.batch_update({"requests": validation_requests})
        
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
            
        # Add notes to term names and controlled vocabulary dropdowns
        note_requests = []
        validation_requests = []
        
        for col_idx in new_cols:
            term_name = sheet_df.iloc[term_name_row, col_idx]
            term_rows = noaa_fields[noaa_fields['term_name'] == term_name]
            if not term_rows.empty:
                term_info = term_rows.iloc[0]
                
                # Add description as note
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
                
                # Add controlled vocabulary dropdown
                if 'controlled_vocabulary_options' in term_info and term_info['controlled_vocabulary_options']:
                    # Parse the controlled vocabulary values
                    cv_values = [v.strip() for v in str(term_info['controlled_vocabulary_options']).split('|') if v.strip()]
                    if cv_values:
                        validation_requests.append({
                            "setDataValidation": {
                                "range": {
                                    "sheetId": worksheet.id,
                                    "startRowIndex": term_name_row + 1,  # Start from the row after term names
                                    "endRowIndex": max(term_name_row + 20, len(updated_data)),  # Ensure we have enough rows
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
        
        # Apply notes
        if note_requests:
            worksheet.spreadsheet.batch_update({"requests": note_requests})
        
        # Apply data validation
        if validation_requests:
            worksheet.spreadsheet.batch_update({"requests": validation_requests})
        
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
        
        # If no analysis runs are specified, or only the placeholder exists, create a single generic analysisMetadata sheet
        if not analysis_runs or "analysisMetadata_<analysis_run_name>" in analysis_runs:
            sheet_name = "analysisMetadata_<analysis_run_name>"
            try:
                # Check if a sheet with this name already exists
                existing_sheet = None
                try:
                    existing_sheet = spreadsheet.worksheet(sheet_name)
                except gspread.exceptions.WorksheetNotFound:
                    pass  # Sheet doesn't exist, which is fine

                if not existing_sheet:
                    worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=100)
                    analysis_worksheets[sheet_name] = worksheet
                else:
                    analysis_worksheets[sheet_name] = existing_sheet

            except gspread.exceptions.APIError as e:
                if "429" in str(e):  # Rate limit error
                    time.sleep(60)
                    worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=100)
                    analysis_worksheets[sheet_name] = worksheet
                else:
                    raise
        else:
            # Create a sheet for each analysis run name
            for analysis_run_name in analysis_runs:
                # If the run name is the placeholder, use it directly. Otherwise, prepend the prefix.
                if analysis_run_name == "analysisMetadata_<analysis_run_name>":
                    sheet_name = analysis_run_name
                else:
                    sheet_name = f"analysisMetadata_{analysis_run_name}"
                try:
                    worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=100)
                    analysis_worksheets[analysis_run_name] = worksheet
                except gspread.exceptions.APIError as e:
                    if "429" in str(e):  # Rate limit error
                        time.sleep(60)
                        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=100)
                        analysis_worksheets[analysis_run_name] = worksheet
                    else:
                        raise
        
        return analysis_worksheets
    
    except Exception as e:
        raise Exception(f"Error creating analysis metadata sheets: {e}")

def add_noaa_fields_to_analysis_metadata(worksheet, noaa_fields, config, analysis_run_name=None):
    """
    Add NOAA fields to analysisMetadata sheet and auto-fill values from config.
    
    Args:
        worksheet (gspread.Worksheet): The analysisMetadata worksheet
        noaa_fields (pandas.DataFrame): DataFrame containing NOAA fields to add
        config (dict): Configuration loaded from NOAA_config.yaml
        analysis_run_name (str, optional): Specific analysis run name for this sheet
    """
    try:
        import pandas as pd
        import gspread
        import gspread_formatting as gsf
        import numpy as np
        import time
        
        # Replace NaN values with empty strings
        noaa_fields = noaa_fields.fillna('')
        
        # Initialize with required headers
        headers = ['requirement_level_code', 'section', 'term_name', 'values']
        
        # Prepare rows to add
        rows = []
        for _, field in noaa_fields.iterrows():
            row = {
                'requirement_level_code': field['requirement_level_code'],
                'section': field['section'],
                'term_name': field['term_name'],
                'values': ''  # Will be filled for auto-fill fields
            }
            
            # Auto-fill values from config for specific fields
            term_name = field['term_name']
            if term_name == 'project_id':
                row['values'] = config['project_id']
            elif term_name == 'analysis_run_name' and analysis_run_name:
                row['values'] = analysis_run_name
            elif term_name == 'assay_name' and analysis_run_name:
                row['values'] = config['analysis_run_name'][analysis_run_name]['assay_name']
                
            rows.append(row)
        
        # Convert to DataFrame for easier handling
        df = pd.DataFrame(rows)
        
        # Prepare all data for a single batch update
        all_data = [headers] + df.values.tolist()
        
        # Define color styles for requirement levels
        color_styles = {
            "M": {"red": 0.89, "green": 0.42, "blue": 0.04},  # #E26B0A - Orange
            "HR": {"red": 1.0, "green": 0.8, "blue": 0.0},    # #FFCC00 - Yellow
            "R": {"red": 1.0, "green": 1.0, "blue": 0.6},     # #FFFF99 - Light yellow
            "O": {"red": 0.8, "green": 1.0, "blue": 0.6}      # #CCFF99 - Light green
        }
        
        # Prepare a single batch request for all operations
        batch_requests = []
        
        # 1. Resize the worksheet
        batch_requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": worksheet.id,
                    "gridProperties": {
                        "rowCount": len(all_data) + 10,  # Add buffer
                        "columnCount": len(headers) + 5   # Add buffer
                    }
                },
                "fields": "gridProperties(rowCount,columnCount)"
            }
        })
        
        # 2. Update all data at once
        batch_requests.append({
            "updateCells": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 0,
                    "endRowIndex": len(all_data),
                    "startColumnIndex": 0,
                    "endColumnIndex": len(headers)
                },
                "rows": [{"values": [{"userEnteredValue": {"stringValue": str(cell)}} for cell in row]} for row in all_data],
                "fields": "userEnteredValue"
            }
        })
        
        # 3. Add header formatting (bold)
        batch_requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": len(headers)
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
        
        # 4. Add formatting for each row
        for i, row in enumerate(rows, start=2):  # Start from row 2 (after headers)
            req_level = row['requirement_level_code']
            
            # Add requirement level color formatting
            if req_level in color_styles:
                batch_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": i-1,
                            "endRowIndex": i,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": color_styles[req_level]
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })
            
            # Add term name bold formatting
            batch_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": i-1,
                        "endRowIndex": i,
                        "startColumnIndex": 2,  # term_name column
                        "endColumnIndex": 3
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
            
            # Add description notes
            description = noaa_fields[noaa_fields['term_name'] == row['term_name']]['description'].iloc[0]
            if description:
                batch_requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": i-1,
                            "endRowIndex": i,
                            "startColumnIndex": 2,
                            "endColumnIndex": 3
                        },
                        "rows": [{
                            "values": [{
                                "note": description
                            }]
                        }],
                        "fields": "note"
                    }
                })
            
            # Add controlled vocabulary dropdowns
            cv_options = noaa_fields[noaa_fields['term_name'] == row['term_name']]['controlled_vocabulary_options'].iloc[0]
            if pd.notna(cv_options) and cv_options:
                values = [v.strip() for v in str(cv_options).split('|') if v.strip()]
                if values:
                    batch_requests.append({
                        "setDataValidation": {
                            "range": {
                                "sheetId": worksheet.id,
                                "startRowIndex": i-1,
                                "endRowIndex": i,
                                "startColumnIndex": 3,  # values column
                                "endColumnIndex": 4
                            },
                            "rule": {
                                "condition": {
                                    "type": "ONE_OF_LIST",
                                    "values": [{"userEnteredValue": v} for v in values]
                                },
                                "showCustomUi": True,
                                "strict": False
                            }
                        }
                    })
        
        # Execute all operations in a single batch request with exponential backoff retry
        max_retries = 5
        retry_count = 0
        wait_time = 2  # Initial wait time in seconds
        
        while retry_count < max_retries:
            try:
                worksheet.spreadsheet.batch_update({"requests": batch_requests})
                break  # Success, exit the retry loop
            except gspread.exceptions.APIError as e:
                if "429" in str(e) and retry_count < max_retries - 1:  # Rate limit error and still have retries
                    retry_count += 1
                    print(f"Rate limit hit. Retrying in {wait_time} seconds (attempt {retry_count}/{max_retries})...")
                    time.sleep(wait_time)
                    wait_time *= 2  # Exponential backoff
                else:
                    raise  # Re-raise if it's not a rate limit error or we're out of retries
        
    except Exception as e:
        raise Exception(f"Error adding NOAA fields to analysisMetadata: {e}")
    
def update_readme_sheet_for_FAIRe2NODE(spreadsheet, config):
    """
    Update the README sheet to reflect the FAIRe2NODE structure.
    
    Args:
        spreadsheet (gspread.Spreadsheet): The Google Spreadsheet object
        config (dict): Configuration loaded from NOAA_config.yaml
        
    Returns:
        None
        
    Raises:
        Exception: If there's an error updating the README sheet
    """
    try:
        # Get the README worksheet
        readme_sheet = spreadsheet.worksheet("README")
        
        # Get all values from the README sheet
        all_values = readme_sheet.get_all_values()
        
        # Find the positions of key sections
        timestamp_section_start = None
        template_params_start = None
        req_levels_start = None
        sheets_section_start = None
        
        for i, row in enumerate(all_values):
            if row and row[0] == 'Modification Timestamp:':
                timestamp_section_start = i
            elif row and row[0] == 'Template parameters:':
                template_params_start = i
            elif row and row[0] == 'Requirement levels:':
                req_levels_start = i
            elif row and row[0] == 'Sheets in this Google sheet:':
                sheets_section_start = i
        
        # Get all worksheet names except README and Drop-down values
        sheet_names = [ws.title for ws in spreadsheet.worksheets() 
                      if ws.title not in ["README", "Drop-down values"]]
        
        # Prepare batch requests for updating content and formatting
        batch_requests = []
        
        # First, resize the worksheet to a reasonable size
        batch_requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": readme_sheet.id,
                    "gridProperties": {
                        "rowCount": 60,  # Set to a reasonable number
                        "columnCount": 10  # Keep the existing column count
                    }
                },
                "fields": "gridProperties.rowCount"
            }
        })
        
        # 1. Change "Modification Timestamp:" to "Sheets in this Google Sheet:"
        if timestamp_section_start is not None:
            batch_requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": timestamp_section_start,
                        "endRowIndex": timestamp_section_start + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "rows": [{"values": [{"userEnteredValue": {"stringValue": "Sheets in this Google Sheet:"}}]}],
                    "fields": "userEnteredValue"
                }
            })
            
            # Format the header row (bold)
            batch_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": timestamp_section_start,
                        "endRowIndex": timestamp_section_start + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
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
            
            # Add the "Sheet Name Timestamp Email" header row
            batch_requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": timestamp_section_start + 1,
                        "endRowIndex": timestamp_section_start + 2,
                        "startColumnIndex": 0,
                        "endColumnIndex": 3
                    },
                    "rows": [{"values": [
                        {"userEnteredValue": {"stringValue": "Sheet Name"}},
                        {"userEnteredValue": {"stringValue": "Timestamp"}},
                        {"userEnteredValue": {"stringValue": "Email"}}
                    ]}],
                    "fields": "userEnteredValue"
                }
            })
            
            # Create rows for the sheet list with empty cells for timestamp and email
            sheet_list_rows = []
            for name in sheet_names:
                sheet_list_rows.append([name, "", ""])
            
            # Add an empty row at the end of the sheet list
            sheet_list_rows.append(["", "", ""])
            
            # Update the sheet list (starting from row after the header)
            sheet_rows_data = []
            for row in sheet_list_rows:
                sheet_row_data = {"values": [
                    {"userEnteredValue": {"stringValue": ""}},
                    {"userEnteredValue": {"stringValue": ""}},
                    {"userEnteredValue": {"stringValue": ""}}
                ]}
                
                # Safely set values for each column
                if len(row) > 0:
                    sheet_row_data["values"][0]["userEnteredValue"]["stringValue"] = str(row[0])
                if len(row) > 1:
                    sheet_row_data["values"][1]["userEnteredValue"]["stringValue"] = str(row[1])
                if len(row) > 2:
                    sheet_row_data["values"][2]["userEnteredValue"]["stringValue"] = str(row[2])
                
                sheet_rows_data.append(sheet_row_data)
            
            # Calculate where Template parameters should start
            template_params_row = timestamp_section_start + 2 + len(sheet_rows_data)
            
            batch_requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": timestamp_section_start + 2,  # +2 to skip header row
                        "endRowIndex": template_params_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": 3  # Include all three columns
                    },
                    "rows": sheet_rows_data,
                    "fields": "userEnteredValue"
                }
            })
            
            # 2. Add "Template Parameters:" after the sheet list
            batch_requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": template_params_row,
                        "endRowIndex": template_params_row + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "rows": [{"values": [{"userEnteredValue": {"stringValue": "Template parameters:"}}]}],
                    "fields": "userEnteredValue"
                }
            })
            
            # Format the header row (bold)
            batch_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": template_params_row,
                        "endRowIndex": template_params_row + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
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
            
            # Find the template parameters content
            template_params_content = []
            if template_params_start is not None:
                for i in range(template_params_start + 1, req_levels_start):
                    if all_values[i] and all_values[i][0]:
                        template_params_content.append([all_values[i][0]])
            
            # Add template parameters content
            if template_params_content:
                batch_requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": readme_sheet.id,
                            "startRowIndex": template_params_row + 1,
                            "endRowIndex": template_params_row + 1 + len(template_params_content),
                            "startColumnIndex": 0,
                            "endColumnIndex": 1
                        },
                        "rows": [{"values": [{"userEnteredValue": {"stringValue": row[0]}}]} for row in template_params_content],
                        "fields": "userEnteredValue"
                    }
                })
                
                # Update where requirement levels should start
                req_levels_row = template_params_row + 1 + len(template_params_content) + 1  # +1 for empty row
            else:
                req_levels_row = template_params_row + 1
            
            # Add empty row after template parameters
            batch_requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": req_levels_row - 1,
                        "endRowIndex": req_levels_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "rows": [{"values": [{"userEnteredValue": {"stringValue": ""}}]}],
                    "fields": "userEnteredValue"
                }
            })
            
            # 3. Add "Requirement levels:" section
            batch_requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": req_levels_row,
                        "endRowIndex": req_levels_row + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "rows": [{"values": [{"userEnteredValue": {"stringValue": "Requirement levels:"}}]}],
                    "fields": "userEnteredValue"
                }
            })
            
            # Format the header row (bold)
            batch_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": req_levels_row,
                        "endRowIndex": req_levels_row + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
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
            
            # Find the requirement levels content
            req_levels_content = []
            if req_levels_start is not None:
                for i in range(req_levels_start + 1, sheets_section_start):
                    if all_values[i] and all_values[i][0]:
                        req_levels_content.append([all_values[i][0]])
            
            # Add requirement levels content
            if req_levels_content:
                batch_requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": readme_sheet.id,
                            "startRowIndex": req_levels_row + 1,
                            "endRowIndex": req_levels_row + 1 + len(req_levels_content),
                            "startColumnIndex": 0,
                            "endColumnIndex": 1
                        },
                        "rows": [{"values": [{"userEnteredValue": {"stringValue": row[0]}}]} for row in req_levels_content],
                        "fields": "userEnteredValue"
                    }
                })
                
                # Add color formatting for requirement level rows
                color_styles = {
                    "M": {"red": 0.89, "green": 0.42, "blue": 0.04},  # #E26B0A - Orange
                    "HR": {"red": 1.0, "green": 0.8, "blue": 0.0},    # #FFCC00 - Yellow
                    "R": {"red": 1.0, "green": 1.0, "blue": 0.6},     # #FFFF99 - Light yellow
                    "O": {"red": 0.8, "green": 1.0, "blue": 0.6}      # #CCFF99 - Light green
                }
                
                # Apply color formatting to each requirement level row
                for i, row in enumerate(req_levels_content):
                    level = row[0].split('=')[0].strip()
                    if level in color_styles:
                        batch_requests.append({
                            "repeatCell": {
                                "range": {
                                    "sheetId": readme_sheet.id,
                                    "startRowIndex": req_levels_row + 1 + i,
                                    "endRowIndex": req_levels_row + 2 + i,
                                    "startColumnIndex": 0,
                                    "endColumnIndex": 1
                                },
                                "cell": {
                                    "userEnteredFormat": {
                                        "backgroundColor": color_styles[level]
                                    }
                                },
                                "fields": "userEnteredFormat.backgroundColor"
                            }
                        })
                
                # Update where instructions should start
                instructions_row = req_levels_row + 1 + len(req_levels_content) + 1  # +1 for empty row
            else:
                instructions_row = req_levels_row + 1
            
            # Add empty row after requirement levels
            batch_requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": instructions_row - 1,
                        "endRowIndex": instructions_row,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "rows": [{"values": [{"userEnteredValue": {"stringValue": ""}}]}],
                    "fields": "userEnteredValue"
                }
            })
            
            # 4. Add "Instructions:" section
            batch_requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": instructions_row,
                        "endRowIndex": instructions_row + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "rows": [{"values": [{"userEnteredValue": {"stringValue": "Instructions:"}}]}],
                    "fields": "userEnteredValue"
                }
            })
            
            # Format the header row (bold)
            batch_requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": instructions_row,
                        "endRowIndex": instructions_row + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
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
            
            # Add instructions content
            instructions = [
                ["1. Enter your data into the projectMetadata, sampleMetadata, and experimentRunMetadata sheets."],
                ["2. If you only have one generic analysisMetadata sheet (meaning you don't have an analysis ready yet) continue to Step 3. If you have analyses ready, skip to Step 4."],
                ["3. Once you do have analyses ready, copy the generic analysisMetadata sheet for as many analyses as you have, and rename them appropriately to analysisMetadata<analysis_run_name>, and fill in the data."],
                ["	- When filling in your data, be careful of the following:"],
                ["	- The 'project_id' MUST be the same as the name of the project_id in the projectMetadata sheet."],
                ["	- The 'analysis_run_name' MUST be the same as the name of the analysis_run_name in the projectMetadata sheet. If you have multiple analyses, seperate each analysis_run_name with a pipe: gomecc4_16s_p1-2_v2024.10_241122 | gomecc4_16s_p3-6_v2024.10_241122. An analysis can only have one analysis_run_name."],
                ["  - The 'assay_name' MUST match one of the assay_names in the projectMetadata sheet. Each assay_name must be seperated by a pipe: ssu16sv4v5-emp |ssu18sv9-emp. An analysis can only have one assay_name."],
                ["4. Fill in the data for you analysisMetadata sheets. Since you specified your analysis_run_names and assay_names in the NOAA_config.yaml file, those fields will auto-fill for you, so you dont have to worry about Step 3."],
                ["5. Optional: For modification history and data validation (Checks for the requirements in Step 3), copy and paste the Google Apps Script from the README into the Google Sheet"],
                ["	- In Google Sheets, go to Extensions > Apps Script > Copy and paste the script > Hit Save"],
                ["6. Ensure all mandatory (M) fields are filled before submission."],
                ["7. Now your data is ready for submission to NODE and edna2obis!"],
                ["8. For each sheet, (except for the README and Drop-down values), download them as a TSV file. This is required for NODE and edna2obis submission."],
                ["	- In Google Sheets, go to File > Download > TSV, for each sheet."],
                ["9. For NODE Submission, go here: https://www.oceandnaexplorer.org/submit"],
                ["10. For edna2obis Submission, go here: https://github.com/aomlomics/edna2obis"],
                ["11. Please don't hesitate to reach out to us with questions or concerns: bayden.willms@noaa.gov"]
            ]
            
            batch_requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": readme_sheet.id,
                        "startRowIndex": instructions_row + 1,
                        "endRowIndex": instructions_row + 1 + len(instructions),
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "rows": [{"values": [{"userEnteredValue": {"stringValue": row[0]}}]} for row in instructions],
                    "fields": "userEnteredValue"
                }
            })
        
        # Apply all batch requests
        if batch_requests:
            readme_sheet.spreadsheet.batch_update({"requests": batch_requests})
            
    except Exception as e:
        raise Exception(f"Error updating README sheet for FAIRe2NODE: {e}")

def show_next_steps_page():
    """
    Opens the next steps page in the default web browser.
    This page shows what to do after successful FAIRe2NODE conversion.
    """
    try:
        # Get the absolute path to the next_steps.html file
        current_dir = os.path.dirname(os.path.abspath(__file__))
        next_steps_path = os.path.join(current_dir, 'next_steps.html')
        
        # Convert the file path to a file URL
        file_url = 'file:///' + next_steps_path.replace('\\', '/')
        
        # Open the page in the default browser
        webbrowser.open(file_url)
        
    except Exception as e:
        print(f"Warning: Could not open next steps page: {e}")