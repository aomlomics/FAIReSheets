"""
FAIRe2NODE - Converts FAIReSheets templates to NODE format.

This script takes a FAIReSheets-generated Google Sheet and modifies it to be compatible
with NODE submission requirements. It removes bioinformatics fields and adds NOAA-specific
fields as needed.
"""

import os
import yaml
import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

from helpers.FAIRe2NODE_helpers import (
    get_bioinformatics_fields,
    remove_bioinfo_fields_from_project_metadata,
    remove_bioinfo_fields_from_experiment_metadata,
)

def FAIRe2NODE(client=None):
    """
    Convert FAIReSheets template to NODE format.
    
    Args:
        client (gspread.Client, optional): Pre-authenticated client. If None, will create one.
    """
    # Load environment variables
    load_dotenv()
    
    # Ensure client is provided
    if client is None:
        raise ValueError("A pre-authenticated client must be provided. Run this function through run.py.")
    
    # Get spreadsheet ID from .env file
    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    if not spreadsheet_id:
        raise ValueError("SPREADSHEET_ID not found in .env file. Please create a .env file with SPREADSHEET_ID=your_sheet_id")
    
    print("Starting FAIRe2NODE conversion...")
    
    # Open the spreadsheet
    spreadsheet = client.open_by_key(spreadsheet_id)
    
    # Load NOAA config
    config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'NOAA_config.yaml')
    try:
        with open(config_path, 'r') as f:
            config = yaml.safe_load(f)
    except Exception as e:
        raise Exception(f"Error reading NOAA config file: {e}")
    
    # Get NOAA checklist path
    noaa_checklist_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'input', 'FAIRe_NOAA_checklist_v1.0.xlsx')
    if not os.path.exists(noaa_checklist_path):
        raise FileNotFoundError(f"NOAA checklist not found at {noaa_checklist_path}")
    
    # Part 1: Remove bioinformatics fields
    print("Part 1: Removing bioinformatics fields...")
    
    # Get bioinformatics fields from NOAA checklist
    bioinfo_fields = get_bioinformatics_fields(noaa_checklist_path)
    print(f"Found {len(bioinfo_fields)} bioinformatics fields to remove")
    
    # Get the worksheets
    try:
        project_metadata = spreadsheet.worksheet("projectMetadata")
        experiment_metadata = spreadsheet.worksheet("experimentRunMetadata")
    except gspread.exceptions.WorksheetNotFound as e:
        raise Exception(f"Required worksheet not found: {e}")
    
    # Remove bioinformatics fields from projectMetadata
    print("Removing bioinformatics fields from projectMetadata...")
    remove_bioinfo_fields_from_project_metadata(project_metadata, bioinfo_fields)
    
    # Remove bioinformatics fields from experimentRunMetadata
    print("Removing bioinformatics fields from experimentRunMetadata...")
    remove_bioinfo_fields_from_experiment_metadata(experiment_metadata, bioinfo_fields)
    
    print("Part 1 completed successfully!")

# Add this code to execute the function when the script is run directly
if __name__ == "__main__":
    from auth import authenticate
    
    print("Starting FAIRe2NODE...")
    try:
        # Authenticate with Google
        client = authenticate()
        
        # Run the conversion
        FAIRe2NODE(client=client)
        
        print("FAIRe2NODE completed successfully!")
    except Exception as e:
        print(f"Error: {e}")
