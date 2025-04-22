"""
FAIRe2NODE - Converts FAIReSheets templates to NODE format.

This script takes a FAIReSheets-generated Google Sheet and modifies it to be compatible
with NODE submission requirements. It removes bioinformatics fields and adds NOAA-specific
fields as needed.
"""

import os
import yaml
import gspread
import gspread_formatting as gsf
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

from helpers.FAIRe2NODE_helpers import (
    get_bioinformatics_fields,
    remove_bioinfo_fields_from_project_metadata,
    remove_bioinfo_fields_from_experiment_metadata,
    get_noaa_fields,
    add_noaa_fields_to_project_metadata,
    add_noaa_fields_to_experiment_metadata,
    add_noaa_fields_to_sample_metadata,
    remove_taxa_sheets,
    create_analysis_metadata_sheets,
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
        sample_metadata = spreadsheet.worksheet("sampleMetadata")
    except gspread.exceptions.WorksheetNotFound as e:
        raise Exception(f"Required worksheet not found: {e}")
    
    # Remove bioinformatics fields from projectMetadata
    print("Removing bioinformatics fields from projectMetadata...")
    remove_bioinfo_fields_from_project_metadata(project_metadata, bioinfo_fields)
    
    # Remove bioinformatics fields from experimentRunMetadata
    print("Removing bioinformatics fields from experimentRunMetadata...")
    remove_bioinfo_fields_from_experiment_metadata(experiment_metadata, bioinfo_fields)
    
    print("Part 1 completed successfully!")
    
    # Part 2: Add NOAA fields to sheets
    print("\nPart 2: Adding NOAA fields to sheets...")
    
    # Define color styles for requirement levels - matching FAIReSheets exactly
    req_col_df = pd.DataFrame({
        'requirement_level': ["M = Mandatory", "HR = Highly recommended", "R = Recommended", "O = Optional"],
        'requirement_level_code': ["M", "HR", "R", "O"],
        'col': ["#E26B0A", "#FFCC00", "#FFFF99", "#CCFF99"]
    })
    
    # Create color styles dictionary
    color_styles = {}
    for _, row in req_col_df.iterrows():
        color_styles[row['requirement_level_code']] = gsf.CellFormat(
            backgroundColor=gsf.Color.fromHex(row['col'])
        )
    
    # Add NOAA project metadata fields
    print("Adding NOAA project metadata fields...")
    noaa_project_fields = get_noaa_fields(noaa_checklist_path, "NOAAprojectMetadata")
    print(f"Found {len(noaa_project_fields)} NOAA project metadata fields to add")
    add_noaa_fields_to_project_metadata(project_metadata, noaa_project_fields)
    
    # Add NOAA sample metadata fields
    print("Adding NOAA sample metadata fields...")
    noaa_sample_fields = get_noaa_fields(noaa_checklist_path, "NOAAsampleMetadata")
    print(f"Found {len(noaa_sample_fields)} NOAA sample metadata fields to add")
    add_noaa_fields_to_sample_metadata(sample_metadata, noaa_sample_fields)
    
    # Add NOAA experiment run metadata fields
    print("Adding NOAA experiment run metadata fields...")
    noaa_experiment_fields = get_noaa_fields(noaa_checklist_path, "NOAAexperimentRunMetadata")
    print(f"Found {len(noaa_experiment_fields)} NOAA experiment run metadata fields to add")
    add_noaa_fields_to_experiment_metadata(experiment_metadata, noaa_experiment_fields)
    
    print("Part 2 completed successfully!")
    
    # Part 3: Remove taxa sheets
    print("\nPart 3: Removing taxa sheets...")
    remove_taxa_sheets(spreadsheet)
    print("Part 3 completed successfully!")
    
    # Part 4: Create analysisMetadata sheets
    print("\nPart 4: Creating analysisMetadata sheets...")
    analysis_worksheets = create_analysis_metadata_sheets(spreadsheet, config)
    print(f"Created {len(analysis_worksheets)} analysisMetadata sheet(s)")
    print("Part 4 completed successfully!")

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
