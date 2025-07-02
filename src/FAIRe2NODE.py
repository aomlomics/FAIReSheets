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
from tqdm import tqdm

from src.helpers.FAIRe2NODE_helpers import (
    get_bioinformatics_fields,
    remove_bioinfo_fields_from_project_metadata,
    remove_bioinfo_fields_from_experiment_metadata,
    get_noaa_fields,
    add_noaa_fields_to_project_metadata,
    add_noaa_fields_to_experiment_metadata,
    add_noaa_fields_to_sample_metadata,
    remove_taxa_sheets,
    create_analysis_metadata_sheets,
    add_noaa_fields_to_analysis_metadata,
    update_readme_sheet_for_FAIRe2NODE,
    show_next_steps_page
)

def FAIRe2NODE(client=None, project_id=None):
    """
    Convert FAIReSheets template to NODE format.
    
    Args:
        client (gspread.Client, optional): Pre-authenticated client. If None, will create one.
        project_id (str, optional): Project ID to use for renaming the spreadsheet.
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
    noaa_checklist_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'input', 'FAIRe_NOAA_checklist_v1.0.2.xlsx')
    if not os.path.exists(noaa_checklist_path):
        raise FileNotFoundError(f"NOAA checklist not found at {noaa_checklist_path}")
    
    # Create a progress bar for the entire process
    pbar = tqdm(total=6, desc="Converting to NODE format", unit="step", position=0, leave=True)
    
    try:
        # Get the worksheets
        project_metadata = spreadsheet.worksheet("projectMetadata")
        experiment_metadata = spreadsheet.worksheet("experimentRunMetadata")
        sample_metadata = spreadsheet.worksheet("sampleMetadata")
        
        # Part 1: Remove bioinformatics fields
        pbar.set_description("Removing bioinformatics fields")
        bioinfo_fields = get_bioinformatics_fields(noaa_checklist_path)
        remove_bioinfo_fields_from_project_metadata(project_metadata, bioinfo_fields)
        remove_bioinfo_fields_from_experiment_metadata(experiment_metadata, bioinfo_fields)
        pbar.update(1)
        
        # Part 2: Add NOAA fields to sheets
        pbar.set_description("Adding NOAA fields")
        noaa_project_fields = get_noaa_fields(noaa_checklist_path, "NOAAprojectMetadata")
        add_noaa_fields_to_project_metadata(project_metadata, noaa_project_fields)
        
        noaa_sample_fields = get_noaa_fields(noaa_checklist_path, "NOAAsampleMetadata")
        add_noaa_fields_to_sample_metadata(sample_metadata, noaa_sample_fields)
        
        noaa_experiment_fields = get_noaa_fields(noaa_checklist_path, "NOAAexperimentRunMetadata")
        add_noaa_fields_to_experiment_metadata(experiment_metadata, noaa_experiment_fields)
        pbar.update(1)
        
        # Part 3: Remove taxa sheets
        pbar.set_description("Removing taxa sheets")
        remove_taxa_sheets(spreadsheet)
        pbar.update(1)
        
        # Part 4: Create analysisMetadata sheets
        pbar.set_description("Creating analysis sheets")
        analysis_worksheets = create_analysis_metadata_sheets(spreadsheet, config)
        pbar.update(1)
        
        # Part 5: Add NOAA analysis metadata fields
        pbar.set_description("Updating metadata")
        noaa_analysis_fields = get_noaa_fields(noaa_checklist_path, "NOAAanalysisMetadata")
        for analysis_run_name, worksheet in analysis_worksheets.items():
            add_noaa_fields_to_analysis_metadata(worksheet, noaa_analysis_fields, config, analysis_run_name)
        update_readme_sheet_for_FAIRe2NODE(spreadsheet, config)
        pbar.update(1)
        
        # Part 6: Rename the spreadsheet
        pbar.set_description("Renaming spreadsheet")
        if project_id:
            new_title = f"FAIRe_NODE_{project_id}"
            spreadsheet.update_title(new_title)
            print(f"\n📝 Spreadsheet renamed to: {new_title}")
        else:
            print("\n⚠️  No project_id provided. Spreadsheet name unchanged.")
        pbar.update(1)
        
        # Close the progress bar
        pbar.close()
        
        # Show the next steps page
        show_next_steps_page()
        
    except Exception as e:
        pbar.close()
        raise Exception(f"Error during conversion: {e}")

# Add this code to execute the function when the script is run directly
if __name__ == "__main__":
    from auth import authenticate
    
    try:
        # Authenticate with Google
        client = authenticate()
        
        # Run the conversion
        FAIRe2NODE(client=client)
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
