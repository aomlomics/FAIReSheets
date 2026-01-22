"""
Helpers for standardizing font formatting in Google Sheets without changing
other formatting (background colors, bold, borders, etc.).
"""

import time

import gspread


def standardize_font_across_spreadsheet(
    spreadsheet,
    font_family="Arial",
    font_size=10,
    max_retries=5,
):
    """
    Set font family + font size for every cell in every worksheet in a spreadsheet.

    IMPORTANT: This uses a Sheets API fields mask to update ONLY:
      - userEnteredFormat.textFormat.fontFamily
      - userEnteredFormat.textFormat.fontSize

    It will not change background colors, boldness, borders, etc.
    """
    worksheets = spreadsheet.worksheets()

    requests = []
    for ws in worksheets:
        # Apply to the entire grid size (row_count x col_count)
        # This matches the user's request: every sheet, every cell.
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": 0,
                        "endRowIndex": ws.row_count,
                        "startColumnIndex": 0,
                        "endColumnIndex": ws.col_count,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "fontFamily": font_family,
                                "fontSize": font_size,
                            }
                        }
                    },
                    "fields": "userEnteredFormat.textFormat.fontFamily,userEnteredFormat.textFormat.fontSize",
                }
            }
        )

    if not requests:
        return

    retry_count = 0
    wait_time = 2
    while retry_count < max_retries:
        try:
            spreadsheet.batch_update({"requests": requests})
            return
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and retry_count < max_retries - 1:
                retry_count += 1
                time.sleep(wait_time)
                wait_time *= 2
            else:
                raise

