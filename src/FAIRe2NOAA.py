"""
FAIRe2NOAA - Converts FAIReSheets templates to NOAA Ocean DNA Explorer input format.

This script takes a FAIReSheets-generated Google Sheet and modifies it to be compatible
with Ocean DNA Explorer submission requirements. It removes bioinformatics fields and
adds NOAA-specific fields as needed.
"""

import os
import yaml
import gspread
import gspread_formatting as gsf
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

from src.helpers.FAIRe2NOAA_helpers import (
    get_bioinformatics_fields,
    remove_bioinfo_fields_from_project_metadata,
    remove_bioinfo_fields_from_experiment_metadata,
    remove_terms_from_experiment_metadata,
    remove_terms_from_sample_metadata,
    get_noaa_fields,
    add_noaa_fields_to_project_metadata,
    add_noaa_fields_to_experiment_metadata,
    add_noaa_fields_to_sample_metadata,
    remove_taxa_sheets,
    create_analysis_metadata_sheets,
    add_noaa_fields_to_analysis_metadata,
    update_readme_sheet_for_FAIRe2NOAA,
    update_noaa_vocab_dropdowns,
    show_next_steps_page
)
from src.helpers.font_standardization import standardize_font_across_spreadsheet

def FAIRe2NOAA(client=None, project_id=None):
    """
    Convert FAIReSheets template to Ocean DNA Explorer format.

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
    
    # Total number of steps
    total_steps = 7
    
    try:
        # Get the worksheets
        project_metadata = spreadsheet.worksheet("projectMetadata")
        experiment_metadata = spreadsheet.worksheet("experimentRunMetadata")
        sample_metadata = spreadsheet.worksheet("sampleMetadata")

        # Part 1: Remove bioinformatics fields
        print(f"Removing bioinformatics fields... (1/{total_steps})")
        bioinfo_fields = get_bioinformatics_fields(noaa_checklist_path)
        remove_bioinfo_fields_from_project_metadata(project_metadata, bioinfo_fields)
        remove_bioinfo_fields_from_experiment_metadata(experiment_metadata, bioinfo_fields)

        # Part 2: Add NOAA fields to sheets
        print(f"Adding NOAA fields to metadata sheets... (2/{total_steps})")
        noaa_project_fields = get_noaa_fields(noaa_checklist_path, "NOAAprojectMetadata")
        add_noaa_fields_to_project_metadata(project_metadata, noaa_project_fields)

        noaa_sample_fields = get_noaa_fields(noaa_checklist_path, "NOAAsampleMetadata")
        add_noaa_fields_to_sample_metadata(sample_metadata, noaa_sample_fields)

        noaa_experiment_fields = get_noaa_fields(noaa_checklist_path, "NOAAexperimentRunMetadata")
        add_noaa_fields_to_experiment_metadata(experiment_metadata, noaa_experiment_fields)

        # NOAA-specific denylist: remove unwanted sampleMetadata terms
        # assay_name is not needed in sampleMetadata for NOAA FAIRe sheets
        remove_terms_from_sample_metadata(
            sample_metadata,
            terms_to_remove=['assay_name']
        )

        # NOAA-specific denylist: remove unwanted experimentRunMetadata terms if present
        remove_terms_from_experiment_metadata(
            experiment_metadata,
            terms_to_remove=[
                'output_read_count',
                'output_otu_num',
                'otu_num_tax_assigned'
            ]
        )

        # Part 3: Remove taxa sheets
        print(f"Removing taxa sheets... (3/{total_steps})")
        remove_taxa_sheets(spreadsheet)

        # Part 4: Create analysisMetadata sheets
        print(f"Creating analysis metadata sheets... (4/{total_steps})")
        analysis_worksheets = create_analysis_metadata_sheets(spreadsheet, config)

        # Part 5: Add NOAA analysis metadata fields
        print(f"Adding NOAA analysis metadata fields... (5/{total_steps})")
        noaa_analysis_fields = get_noaa_fields(noaa_checklist_path, "NOAAanalysisMetadata")
        for analysis_run_name, worksheet in analysis_worksheets.items():
            add_noaa_fields_to_analysis_metadata(worksheet, noaa_analysis_fields, config, analysis_run_name)
        update_readme_sheet_for_FAIRe2NOAA(spreadsheet, config)
        
        # Part 6: Update dropdown values with NOAA-specific vocabulary
        print(f"Updating dropdown values with NOAA vocabulary... (6/{total_steps})")
        update_noaa_vocab_dropdowns(spreadsheet, noaa_checklist_path)
        
        # Part 7: Rename the spreadsheet
        if project_id:
            print(f"Renaming spreadsheet... (7/{total_steps})")
            new_title = f"FAIRe-NOAA_{project_id}"
            spreadsheet.update_title(new_title)
        else:
            pass

        # Standardize font across the final spreadsheet (every sheet, every cell).
        # This updates ONLY font family + size via a fields mask, leaving background/bold/etc. intact.
        standardize_font_across_spreadsheet(spreadsheet)

        # Show the next steps page
        show_next_steps_page()

    except Exception as e:
        raise Exception(f"Error during conversion: {e}")


# Add this code to execute the function when the script is run directly
if __name__ == "__main__":
    from auth import authenticate

    try:
        # Authenticate with Google
        client = authenticate()

        # Run the conversion
        FAIRe2NOAA(client=client)
        
    except Exception as e:
        print(f"\nError: {e}")
