"""
Module for creating the Drop-down values sheet in FAIReSheets.
"""

import gspread_formatting as gsf

def create_dropdown_sheet(worksheet, vocab_df, assay_type, assay_name):
    """Create the Drop-down values sheet with controlled vocabulary options."""
    
    # Simply copy the Drop-down values sheet from the template
    if not vocab_df.empty:
        # Replace NaN values with empty strings to avoid JSON errors
        vocab_df_clean = vocab_df.fillna('')

        # Convert DataFrame to list of lists for gspread
        data = [vocab_df_clean.columns.tolist()] + vocab_df_clean.values.tolist()
        
        # Update the worksheet
        worksheet.update("A1", data)
        
        # Format header row
        header_format = gsf.CellFormat(textFormat=gsf.TextFormat(bold=True))
        gsf.format_cell_range(worksheet, "1:1", header_format)
        
        print("Created Drop-down values sheet") 