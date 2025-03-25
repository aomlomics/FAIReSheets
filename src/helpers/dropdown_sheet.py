"""
Module for creating the drop-down values sheet in FAIReSheets.
"""

def create_dropdown_sheet(worksheet, vocab_df, assay_type, assay_name):
    """Create and populate a sheet with all dropdown values."""
    
    # Replace NaN values with empty strings
    vocab_df = vocab_df.fillna('')
    
    # Convert to list of lists for gspread
    data = [vocab_df.columns.tolist()] + vocab_df.values.tolist()
    
    # Resize the worksheet to accommodate the data
    rows_needed = len(data) + 5  # Add buffer
    cols_needed = len(data[0]) + 2  # Add buffer
    worksheet.resize(rows=rows_needed, cols=cols_needed)
    
    # Update the worksheet
    worksheet.update("A1", data)
    
    # Format the headers
    worksheet.format("1:1", {
        "textFormat": {
            "bold": True
        }
    }) 