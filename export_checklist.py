"""
Export the FAIRe checklist as a CSV you can import onto the checklist tab
of an existing Google Sheet.

Usage (from the project root):
    python export_checklist.py FAIRe_NOAA_checklist_v1.0.3.xlsx
    python export_checklist.py FAIRe_checklist_v1.0.3.xlsx
"""

import argparse
import os
import sys

from src.helpers.checklist_sheet import export_checklist_csv


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Export a checklist .xlsx from input/ as input/checklist.csv "
            "for File > Import onto an existing Google Sheet checklist tab."
        )
    )
    parser.add_argument(
        'checklist_name',
        help='Filename of the checklist in input/ (paste the name, including .xlsx)',
    )
    args = parser.parse_args()

    repo_root = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(repo_root, 'input')
    xlsx_path = os.path.join(input_dir, os.path.basename(args.checklist_name))

    if not os.path.isfile(xlsx_path):
        sys.exit(f'File not found in input/: {os.path.basename(args.checklist_name)}')

    output_path = os.path.join(input_dir, 'checklist.csv')
    export_checklist_csv(xlsx_path, output_path)

    print(f'Source: {xlsx_path}')
    print(f'Wrote:  {output_path}')
    print(
        'In your Google Sheet, open the checklist tab, then '
        'File > Import > Upload this CSV > Replace current sheet.'
    )


if __name__ == '__main__':
    main()
