"""
FAIR eDNA Template Generator for Google Sheets (FAIReSheets)

This script generates FAIR eDNA data templates in Google Sheets, based on 
sample types, assay type, and requirement levels of your choice.

Instructions:

Step 1: Save the input files in the working directory
    - FAIRe_checklist_v1.0.xlsx
    - FAIRe_checklist_v1.0_FULLtemplate.xlsx

Step 2: Create a Google Sheet and note its ID
    - Create an empty Google Sheet
    - Get the ID from the URL (the long string between /d/ and /edit in the URL)
    - Add this ID to your .env file as SPREADSHEET_ID=your_sheet_id

Step 3: Ensure you have the service account JSON file in your directory
    - The file should be named fairesheets-609bb159302b.json or specified in your .env
    - Add SERVICE_ACCOUNT_FILE=your_file_name.json to .env if using a different file

Step 4: Run the FAIReSheets function with the required arguments
"""

import os
import pandas as pd
import numpy as np
import gspread
import gspread_formatting as gsf
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import time

# Import tqdm for progress bar
try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False
    print("Warning: tqdm package not found. Install with 'pip install tqdm' for progress bar visualization.")

# Import functions from separate modules
from src.helpers.readme_sheet import create_readme_sheet
from src.helpers.dropdown_sheet import create_dropdown_sheet
from src.helpers.project_metadata_sheet import create_project_metadata_sheet
from src.helpers.sample_metadata_sheet import create_sample_metadata_sheet
from src.helpers.experiment_metadata_sheet import create_experiment_metadata_sheet
from src.helpers.taxa_sheets import create_taxa_sheets
from src.helpers.targeted_sheets import create_targeted_sheets

def FAIReSheets(req_lev=['M', 'HR', 'R', 'O'],
                sample_type=None, 
                assay_type=None, 
                project_id=None, 
                assay_name=None, 
                projectMetadata_user=None,
                sampleMetadata_user=None,
                experimentRunMetadata_user=None,
                input_dir=None,
                client=None):
    """
    Generate FAIR eDNA data templates in Google Sheets
    
    Parameters:
    -----------
    req_lev : list
        Requirement level(s) of fields to include. Select from "M", "HR", "R", "O"
        (Mandatory, Highly recommended, Recommended, Optional). Default: all levels
    
    sample_type : list or str
        Sample type(s). Select from "Water", "Soil", "Sediment", "Air", 
        "HostAssociated", "MicrobialMatBiofilm", "SymbiontAssociated", or "other".
        "other" will include all sample-type-specific fields.
    
    assay_type : str
        Approach applied to detect taxon/taxa. Either "targeted" (e.g., q/d PCR) 
        or "metabarcoding".
    
    project_id : str
        A brief project identifier with no spaces or special characters.
    
    assay_name : list or str
        Assay name(s) with no spaces or special characters.
    
    projectMetadata_user : list, optional
        User-defined fields not listed in the FAIR eDNA metadata checklist.
        These fields will be appended to projectMetadata.
    
    sampleMetadata_user : list, optional
        User-defined fields not listed in the FAIR eDNA metadata checklist.
        These fields will be appended to sampleMetadata.
        
    experimentRunMetadata_user : list, optional
        User-defined fields not listed in the FAIR eDNA metadata checklist.
        These fields will be appended to experimentRunMetadata.
    
    input_dir : str, optional
        Directory containing the input files. If not provided, the current
        working directory is used.
        
    client : gspread.Client, required
        Pre-authenticated gspread client from OAuth authentication.
    
    Returns:
    --------
    None
        Creates/updates a Google Sheet with the specified template
    """

    # Convert single strings to lists for consistency
    if isinstance(sample_type, str):
        sample_type = [sample_type]
    if isinstance(assay_name, str):
        assay_name = [assay_name]
    if isinstance(projectMetadata_user, str):
        projectMetadata_user = [projectMetadata_user]
    if isinstance(sampleMetadata_user, str):
        sampleMetadata_user = [sampleMetadata_user]
    if isinstance(experimentRunMetadata_user, str):
        experimentRunMetadata_user = [experimentRunMetadata_user]

    # Initialize empty lists if None was provided
    if projectMetadata_user is None:
        projectMetadata_user = []
    if sampleMetadata_user is None:
        sampleMetadata_user = []
    if experimentRunMetadata_user is None:
        experimentRunMetadata_user = []
        
    # Load environment variables
    load_dotenv()
    
    # Ensure client is provided
    if client is None:
        raise ValueError("A pre-authenticated client must be provided. Run this function through run.py.")
    
    # Get spreadsheet ID from .env file
    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    if not spreadsheet_id:
        raise ValueError("SPREADSHEET_ID not found in .env file. Please create a .env file with SPREADSHEET_ID=your_sheet_id")
    
    print("Starting template generation...")
    
    # Open the spreadsheet
    spreadsheet = client.open_by_key(spreadsheet_id)
    
    # Update the spreadsheet title to include the project_id
    spreadsheet.update_title(f"FAIRe_{project_id}")
    
    # Set input files
    FAIRe_checklist_ver = 'v1.0'
    input_file_name = f'FAIRe_checklist_{FAIRe_checklist_ver}.xlsx'
    sheet_name = 'checklist'
    
    # Set the file paths correctly
    if input_dir:
        input_file_path = os.path.join(input_dir, input_file_name)
    else:
        # Look in the 'input' directory by default
        input_file_path = os.path.join('input', input_file_name)
    
    # Read input checklist
    try:
        input_df = pd.read_excel(input_file_path, sheet_name=sheet_name)
    except FileNotFoundError:
        raise FileNotFoundError(f"Could not find input file {input_file_path}. Please ensure it is in the specified directory.")
    
    # Full template file name
    full_temp_file_name = f'FAIRe_checklist_{FAIRe_checklist_ver}_FULLtemplate.xlsx'
    if input_dir:
        full_temp_file_path = os.path.join(input_dir, full_temp_file_name)
    else:
        # Look in the 'input' directory by default
        full_temp_file_path = os.path.join('input', full_temp_file_name)
        
    try:
        # Read Excel file using pandas instead of openpyxl directly
        full_template_df = pd.read_excel(full_temp_file_path, sheet_name=None, engine='openpyxl')
    except Exception as e:
        raise Exception(f"Error reading Excel file with pandas: {e}")
    
    # Colors for requirement levels - matching the R script exactly
    req_col_df = pd.DataFrame({
        'requirement_level': ["M = Mandatory", "HR = Highly recommended", "R = Recommended", "O = Optional"],
        'requirement_level_code': ["M", "HR", "R", "O"],
        'col': ["#E26B0A", "#FFCC00", "#FFFF99", "#CCFF99"]
    })
    
    # Define color styles
    color_styles = {}
    for _, row in req_col_df.iterrows():
        color_styles[row['requirement_level_code']] = gsf.CellFormat(
            backgroundColor=gsf.Color.fromHex(row['col'])
        )
    
    # Create or clear sheets
    # First create a list of all sheets we'll need (excluding README which will use Sheet1)
    sheet_names = ["projectMetadata", "sampleMetadata", "Drop-down values"]
    
    # Add assay-type specific sheets
    if assay_type == 'metabarcoding':
        sheet_names.extend(["experimentRunMetadata", "taxaRaw", "taxaFinal"])
    elif assay_type == 'targeted':
        sheet_names.extend(["stdData", "eLowQuantData", "ampData"])
    
    # Get existing worksheet names
    existing_sheets = [ws.title for ws in spreadsheet.worksheets()]
    
    # Create a list of all operations to perform for progress tracking
    operations = []
    operations.append("Setup")  # Initial setup
    operations.append("README")  # README sheet
    
    # Add all sheet names
    for sheet_name in sheet_names:
        operations.append(f"{sheet_name}")
    
    # Initialize progress bar
    if TQDM_AVAILABLE:
        pbar = tqdm(total=len(operations), desc="Initializing...", unit="sheet")
    
    # Delete existing sheets if they match our names
    for sheet_name in existing_sheets:
        if sheet_name in sheet_names:
            worksheet = spreadsheet.worksheet(sheet_name)
            spreadsheet.del_worksheet(worksheet)
    
    # Create new sheets
    worksheets = {}
    for sheet_name in sheet_names:
        # Create with more rows and columns by default
        worksheets[sheet_name] = spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=100)
    
    # Use Sheet1 as README sheet
    try:
        readme_sheet = spreadsheet.worksheet("Sheet1")
        # Rename Sheet1 to README
        readme_sheet.update_title("README")
        worksheets["README"] = readme_sheet
    except gspread.exceptions.WorksheetNotFound:
        # If Sheet1 doesn't exist for some reason, create README sheet
        worksheets["README"] = spreadsheet.add_worksheet(title="README", rows=200, cols=100)
    
    # Update progress bar for setup
    if TQDM_AVAILABLE:
        pbar.update(1)
        pbar.set_description(f"Setting up sheets [1/{len(operations)}]")
    else:
        print("Sheet setup complete (1/{})".format(len(operations)))
    
    # Create README sheet
    if TQDM_AVAILABLE:
        pbar.set_description("Creating README sheet...")
        
    create_readme_sheet(
        worksheet=worksheets["README"],
        input_file_name=input_file_name,
        req_lev=req_lev,
        sample_type=sample_type,
        assay_type=assay_type,
        project_id=project_id,
        assay_name=assay_name,
        projectMetadata_user=projectMetadata_user,
        sampleMetadata_user=sampleMetadata_user,
        color_styles=color_styles,
        FAIRe_checklist_ver=FAIRe_checklist_ver
    )
    
    # Update progress bar for README
    if TQDM_AVAILABLE:
        pbar.update(1)
        pbar.set_description(f"README sheet created [2/{len(operations)}]")
    else:
        print("README sheet created (2/{})".format(len(operations)))
    
    # Read vocabulary data from the full template
    vocab_df = pd.read_excel(full_temp_file_path, sheet_name='Drop-down values')
    
    # Create Drop-down values sheet
    if TQDM_AVAILABLE:
        pbar.set_description("Creating Drop-down values sheet...")
        
    create_dropdown_sheet(
        worksheet=worksheets["Drop-down values"],
        vocab_df=vocab_df,
        assay_type=assay_type,
        assay_name=assay_name
    )
    
    # Update progress bar for Drop-down values
    if TQDM_AVAILABLE:
        pbar.update(1)
        pbar.set_description(f"Drop-down values created [3/{len(operations)}]")
    else:
        print("Drop-down values sheet created (3/{})".format(len(operations)))
    
    # ----- Project Metadata Sheet (Known to be slow) -----
    if TQDM_AVAILABLE:
        pbar.set_description("Creating Project Metadata sheet...")
    
    # Create projectMetadata sheet
    create_project_metadata_sheet(
        worksheet=worksheets["projectMetadata"],
        full_temp_file_name=full_temp_file_path,
        input_df=input_df,
        req_lev=req_lev,
        assay_type=assay_type,
        project_id=project_id,
        assay_name=assay_name,
        projectMetadata_user=projectMetadata_user,
        color_styles=color_styles,
        vocab_df=vocab_df,
        FAIRe_checklist_ver=FAIRe_checklist_ver
    )
    
    # Update progress bar for projectMetadata
    if TQDM_AVAILABLE:
        pbar.update(1)
        pbar.set_description(f"Project Metadata created [4/{len(operations)}]")
    else:
        print("Project Metadata sheet created (4/{})".format(len(operations)))
    
    # ----- Sample Metadata Sheet (Known to be slow) -----
    if TQDM_AVAILABLE:
        pbar.set_description("Creating Sample Metadata sheet...")
    
    # Create sampleMetadata sheet
    create_sample_metadata_sheet(
        worksheet=worksheets["sampleMetadata"],
        full_temp_file_name=full_temp_file_path,
        input_df=input_df,
        req_lev=req_lev,
        sample_type=sample_type,
        assay_type=assay_type,
        assay_name=assay_name,
        sampleMetadata_user=sampleMetadata_user,
        color_styles=color_styles,
        vocab_df=vocab_df
    )
    
    # Update progress bar for sampleMetadata
    if TQDM_AVAILABLE:
        pbar.update(1)
        pbar.set_description(f"Sample Metadata created [5/{len(operations)}]")
    else:
        print("Sample Metadata sheet created (5/{})".format(len(operations)))
    
    # Create assay-type specific sheets
    if assay_type == 'metabarcoding':
        # ----- Experiment Metadata Sheet (Can be slow) -----
        if TQDM_AVAILABLE:
            pbar.set_description("Creating Experiment Run Metadata sheet...")
        
        # Use the specialized function for experimentRunMetadata
        create_experiment_metadata_sheet(
            worksheet=worksheets["experimentRunMetadata"],
            full_temp_file_name=full_temp_file_path,
            input_df=input_df,
            req_lev=req_lev,
            color_styles=color_styles,
            vocab_df=vocab_df,
            experimentRunMetadata_user=experimentRunMetadata_user
        )
        
        # Update progress bar for experimentRunMetadata
        if TQDM_AVAILABLE:
            pbar.update(1)
            pbar.set_description(f"Experiment Run Metadata created [6/{len(operations)}]")
        else:
            print("Experiment Run Metadata sheet created (6/{})".format(len(operations)))
        
        # Use the specialized function for taxa sheets - process both sheets at once
        taxa_sheet_names = ["taxaRaw", "taxaFinal"]
        for sheet_name in taxa_sheet_names:
            # Taxa sheets can also be slow
            if TQDM_AVAILABLE:
                pbar.set_description(f"Creating {sheet_name} sheet...")
            
            create_taxa_sheets(
                worksheet=worksheets[sheet_name],
                sheet_name=sheet_name,
                full_temp_file_name=full_temp_file_path,
                input_df=input_df,
                req_lev=req_lev,
                color_styles=color_styles,
                vocab_df=vocab_df
            )
            
            # Update progress bar for each taxa sheet
            if TQDM_AVAILABLE:
                pbar.update(1)
                current_operation = operations.index(sheet_name) + 1
                pbar.set_description(f"{sheet_name} created [{current_operation}/{len(operations)}]")
            else:
                current_operation = operations.index(sheet_name) + 1
                print(f"{sheet_name} sheet created ({current_operation}/{len(operations)})")
    
    elif assay_type == 'targeted':
        # Get the targeted sheet names
        targeted_sheet_names = ["stdData", "eLowQuantData", "ampData"]
        
        # Targeted sheets can be slow
        if TQDM_AVAILABLE:
            pbar.set_description("Creating targeted assay sheets...")
        
        create_targeted_sheets(
            worksheets=worksheets,
            sheet_names=targeted_sheet_names,
            full_temp_file_name=full_temp_file_path,
            input_df=input_df,
            req_lev=req_lev,
            color_styles=color_styles,
            vocab_df=vocab_df,
            project_id=project_id,
            assay_name=assay_name
        )
        
        # Update progress bar for all three targeted sheets
        for sheet_name in targeted_sheet_names:
            if TQDM_AVAILABLE:
                pbar.update(1)
                current_operation = operations.index(sheet_name) + 1
                pbar.set_description(f"{sheet_name} created [{current_operation}/{len(operations)}]")
            else:
                current_operation = operations.index(sheet_name) + 1
                print(f"{sheet_name} sheet created ({current_operation}/{len(operations)})")
        
    # Wait a moment to ensure all operations are complete
    time.sleep(1)
    
    # Close the progress bar if it exists
    if TQDM_AVAILABLE:
        pbar.set_description("Template generation complete!")
        pbar.close()
    
    print("\nTemplate generation complete!")