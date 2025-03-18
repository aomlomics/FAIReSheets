import os
import sys
import yaml

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
        # Load configuration from config.yaml
        config_path = os.path.join(current_dir, "config.yaml")
        
        if not os.path.exists(config_path):
            print(f"ERROR: Config file not found at {config_path}")
            sys.exit(1)
            
        with open(config_path, 'r') as config_file:
            config = yaml.safe_load(config_file)
            
        print(f"Loaded configuration from {config_path}")
        
        # Check if the Excel files exist
        FAIRe_checklist_ver = 'v1.0'
        input_file_name = f'FAIRe_checklist_{FAIRe_checklist_ver}.xlsx'
        full_temp_file_name = f'FAIRe_checklist_{FAIRe_checklist_ver}_FULLtemplate.xlsx'
        
        print(f"Checking for input files in: {current_dir}")
        print(f"Looking for: {input_file_name} and {full_temp_file_name}")
        
        if not os.path.exists(input_file_name):
            print(f"ERROR: Could not find {input_file_name}")
            sys.exit(1)
        else:
            print(f"Found {input_file_name}")
            
        if not os.path.exists(full_temp_file_name):
            print(f"ERROR: Could not find {full_temp_file_name}")
            sys.exit(1)
        else:
            print(f"Found {full_temp_file_name}")
            
            # Try to load the template with pandas
            try:
                print("Testing Excel file reading...")
                import pandas as pd
                sheets = pd.ExcelFile(full_temp_file_name).sheet_names
                print(f"Excel file contains sheets: {sheets}")
                print("Excel file appears to be valid")
            except Exception as e:
                print(f"Error reading Excel file: {e}")
                sys.exit(1)
        
        # Call the function with parameters from config.yaml
        print("\nRunning FAIReSheets function...")
        FAIReSheets(
            req_lev=config.get('req_lev', ['M', 'HR', 'R', 'O']),
            sample_type=config.get('sample_type', ['Water']),
            assay_type=config.get('assay_type', 'metabarcoding'),
            project_id=config.get('project_id', 'default_project'),
            assay_name=config.get('assay_name', ['default_assay']),
            projectMetadata_user=config.get('projectMetadata_user', None),
            sampleMetadata_user=config.get('sampleMetadata_user', None)
        )
        
        print("\nFAIReSheets completed successfully!")
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc() 