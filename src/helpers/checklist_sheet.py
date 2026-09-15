"""
Module for creating the checklist sheet in FAIReSheets.
"""

import glob
import os
import re

import pandas as pd


def default_input_dir():
    """Path to the repo input/ folder."""
    helpers_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(os.path.dirname(os.path.dirname(helpers_dir)), 'input')


def _version_key(filename):
    match = re.search(r'v(\d+(?:\.\d+)*)', filename, re.IGNORECASE)
    if not match:
        return (0,)
    return tuple(int(part) for part in match.group(1).split('.'))


def _pick_latest(paths):
    paths = [p for p in paths if not os.path.basename(p).startswith('~$')]
    paths.sort(key=lambda p: (_version_key(os.path.basename(p)), os.path.getmtime(p)))
    return paths[-1]


def find_checklist_xlsx(input_dir, use_noaa=False):
    """
    Find a checklist .xlsx in input_dir. Version is taken from the filename,
    not from code.
    """
    input_dir = input_dir or default_input_dir()

    if use_noaa:
        matches = glob.glob(os.path.join(input_dir, 'FAIRe_NOAA_checklist_*.xlsx'))
        expected = 'FAIRe_NOAA_checklist_*.xlsx'
    else:
        matches = [
            path for path in glob.glob(os.path.join(input_dir, 'FAIRe_checklist_*.xlsx'))
            if 'FULLtemplate' not in os.path.basename(path)
            and 'NOAA' not in os.path.basename(path)
        ]
        expected = 'FAIRe_checklist_*.xlsx (not the FULLtemplate)'

    if not matches:
        raise FileNotFoundError(
            f"No checklist file matching {expected} in {input_dir}. "
            "Put the official checklist .xlsx in the input/ folder."
        )

    return _pick_latest(matches)


def find_fulltemplate_xlsx(input_dir):
    """Find the FULLtemplate .xlsx in input_dir. Version is taken from the filename."""
    input_dir = input_dir or default_input_dir()
    matches = glob.glob(os.path.join(input_dir, 'FAIRe_checklist_*FULLtemplate*.xlsx'))
    if not matches:
        raise FileNotFoundError(
            f"No FULLtemplate file matching FAIRe_checklist_*FULLtemplate*.xlsx in {input_dir}."
        )
    return _pick_latest(matches)


def prepare_checklist_df(checklist_df):
    """Normalize checklist values the same way for Google Sheets and local CSV."""
    checklist_df = checklist_df.fillna('')
    return checklist_df.astype(str).replace(['nan', 'NaT', 'None'], '')


def export_checklist_csv(xlsx_path, output_path):
    """Write the checklist sheet as a CSV that Google Sheets can import cleanly."""
    checklist_df = prepare_checklist_df(pd.read_excel(xlsx_path, sheet_name='checklist'))
    output_dir = os.path.dirname(os.path.abspath(output_path))
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    checklist_df.to_csv(output_path, index=False, encoding='utf-8-sig')
    return output_path


def create_checklist_sheet(worksheet, checklist_df):
    """Create and populate a sheet with the full FAIRe checklist."""

    checklist_df = prepare_checklist_df(checklist_df)

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
