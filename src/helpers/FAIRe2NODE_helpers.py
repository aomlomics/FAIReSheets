"""
Module for FAIRe2NODE helper functions.
These functions support the conversion of FAIReSheets templates to NODE format.
"""

import pandas as pd
import gspread
import time

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
        
        # Get all fields where Section is 'Bioinformatics'
        bioinfo_fields = input_df[input_df['Section'] == 'Bioinformatics']['term_name'].tolist()
        
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