"""
Module for creating other sheet types in FAIReSheets.
"""

import pandas as pd

def create_other_sheets(worksheets, sheet_names, full_temp_file_name, input_df, req_lev, color_styles, vocab_df):
    """Create and format other sheets based on assay type."""
    
    for sheet_name in sheet_names:
        worksheet = worksheets[sheet_name]
        
        # Read the template using pandas instead of openpyxl
        sheet_df = pd.read_excel(full_temp_file_name, sheet_name=sheet_name, header=None)
        
        # Replace NaN values with empty strings to avoid JSON errors
        sheet_df = sheet_df.fillna('')
        
        # Convert to list of lists for gspread
        data = sheet_df.values.tolist()
        
        # Resize worksheet to accommodate all data (add some buffer)
        rows_needed = len(data) + 20  # Add buffer
        cols_needed = len(data[0]) + 10 if data else 50  # Add buffer
        worksheet.resize(rows=rows_needed, cols=cols_needed)
        
        # Write to Google Sheets
        worksheet.update("A1", data) 