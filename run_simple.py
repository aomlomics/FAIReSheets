import os
import sys

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Your Google Sheet ID
spreadsheet_id = "1Tfnhkm2IC8LSheYPFmt8JzUiBm7talV5nebSdMrhVXM"

# Set environment variables directly in the script
os.environ["SPREADSHEET_ID"] = spreadsheet_id
os.environ["SERVICE_ACCOUNT_FILE"] = os.path.join(current_dir, "fairesheets-609bb159302b.json")

print(f"Using credentials file: {os.environ['SERVICE_ACCOUNT_FILE']}")
print(f"Using spreadsheet ID: {spreadsheet_id}")

# Import the function directly from the file in the same directory
from FAIReSheets import FAIReSheets

if __name__ == "__main__":
    try:
        # Add a fix for the Excel file issue
        import pandas as pd
        from openpyxl import load_workbook
        
        # Check if the Excel files exist
        FAIRe_checklist_ver = 'v1.0'
        input_file_name = f'FAIRe_checklist_{FAIRe_checklist_ver}.xlsx'
        full_temp_file_name = f'FAIRe_checklist_{FAIRe_checklist_ver}_FULLtemplate.xlsx'
        
        print(f"Checking for input files in: {current_dir}")
        print(f"Looking for: {input_file_name} and {full_temp_file_name}")
        
        if not os.path.exists(input_file_name):
            print(f"ERROR: Could not find {input_file_name}")
        else:
            print(f"Found {input_file_name}")
            
        if not os.path.exists(full_temp_file_name):
            print(f"ERROR: Could not find {full_temp_file_name}")
        else:
            print(f"Found {full_temp_file_name}")
            
            # Try to load the template with pandas instead of openpyxl
            try:
                print("Testing Excel file reading...")
                sheets = pd.ExcelFile(full_temp_file_name).sheet_names
                print(f"Excel file contains sheets: {sheets}")
                print("Excel file appears to be valid")
            except Exception as e:
                print(f"Error reading Excel file: {e}")
        
        # Call the function with your parameters
        print("\nRunning FAIReSheets function...")
        FAIReSheets(
            req_lev=['M', 'HR', 'R', 'O'],
            sample_type=['Water'],
            assay_type='metabarcoding',
            project_id='gomecc_4_GSHEET_test',
            assay_name=['18S_TEST', '16S_TEST']
        )
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc() 