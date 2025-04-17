"""
Module for FAIRe2NODE helper functions.
These functions support the conversion of FAIReSheets templates to NODE format.
"""

import pandas as pd
import gspread
import time
import os

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
        
        if not rows_to_delete:
            return
            
        # Prepare batch delete request
        # Note: We need to delete from bottom to top to maintain correct indices
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