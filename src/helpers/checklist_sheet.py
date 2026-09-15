"""
Module for creating the checklist sheet in FAIReSheets.
"""


def create_checklist_sheet(worksheet, checklist_df):
    """Create and populate a sheet with the full FAIRe checklist."""

    # Replace NaN values with empty strings
    checklist_df = checklist_df.fillna('')

    # Stringify so dates/numpy types serialize for the Sheets API
    checklist_df = checklist_df.astype(str).replace(['nan', 'NaT', 'None'], '')

    # Convert to list of lists for gspread
    data = [checklist_df.columns.tolist()] + checklist_df.values.tolist()

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
