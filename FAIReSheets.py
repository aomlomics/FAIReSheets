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
import re
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
import gspread
import gspread_formatting as gsf
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import time

def FAIReSheets(req_lev=['M', 'HR', 'R', 'O'],
                sample_type=None, 
                assay_type=None, 
                project_id=None, 
                assay_name=None, 
                projectMetadata_user=None,
                sampleMetadata_user=None):
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

    # Initialize empty lists if None was provided
    if projectMetadata_user is None:
        projectMetadata_user = []
    if sampleMetadata_user is None:
        sampleMetadata_user = []
        
    # Load environment variables
    load_dotenv()
    
    # Authentication setup
    json_key_file = os.getenv("SERVICE_ACCOUNT_FILE", "fairesheets-609bb159302b.json")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = Credentials.from_service_account_file(json_key_file, scopes=scope)
    client = gspread.authorize(credentials)
    
    # Get spreadsheet ID from .env file
    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    if not spreadsheet_id:
        raise ValueError("SPREADSHEET_ID not found in .env file. Please create a .env file with SPREADSHEET_ID=your_sheet_id")
    
    # Open the spreadsheet
    spreadsheet = client.open_by_key(spreadsheet_id)
    
    # Update the spreadsheet title to include the project_id
    spreadsheet.update_title(f"FAIRe_{project_id}")
    
    # Set input files
    FAIRe_checklist_ver = 'v1.0'
    input_file_name = f'FAIRe_checklist_{FAIRe_checklist_ver}.xlsx'
    sheet_name = 'checklist'
    
    # Read input checklist
    try:
        input_df = pd.read_excel(input_file_name, sheet_name=sheet_name)
    except FileNotFoundError:
        raise FileNotFoundError(f"Could not find input file {input_file_name}. Please ensure it is in the working directory.")
    
    # Full template file name
    full_temp_file_name = f'FAIRe_checklist_{FAIRe_checklist_ver}_FULLtemplate.xlsx'
    try:
        # Read Excel file using pandas instead of openpyxl directly
        full_template_df = pd.read_excel(full_temp_file_name, sheet_name=None, engine='openpyxl')
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
    # First create a list of all sheets we'll need
    sheet_names = ["README", "projectMetadata", "sampleMetadata", "Drop-down values"]
    
    # Add assay-type specific sheets
    if assay_type == 'metabarcoding':
        sheet_names.extend(["experimentRunMetadata", "taxaRaw", "taxaFinal"])
    elif assay_type == 'targeted':
        sheet_names.extend(["stdData", "eLowQuantData", "ampData"])
    
    # Get existing worksheet names
    existing_sheets = [ws.title for ws in spreadsheet.worksheets()]
    
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
    
    # Create README sheet
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
    
    # Read vocabulary data from the full template
    vocab_df = pd.read_excel(full_temp_file_name, sheet_name='Drop-down values')
    
    # Create Drop-down values sheet
    create_dropdown_sheet(
        worksheet=worksheets["Drop-down values"],
        vocab_df=vocab_df,
        assay_type=assay_type,
        assay_name=assay_name
    )
    
    # Create projectMetadata sheet
    create_project_metadata_sheet(
        worksheet=worksheets["projectMetadata"],
        full_temp_file_name=full_temp_file_name,
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
    
    # Create sampleMetadata sheet
    create_sample_metadata_sheet(
        worksheet=worksheets["sampleMetadata"],
        full_temp_file_name=full_temp_file_name,
        input_df=input_df,
        req_lev=req_lev,
        sample_type=sample_type,
        assay_type=assay_type,
        assay_name=assay_name,
        sampleMetadata_user=sampleMetadata_user,
        color_styles=color_styles,
        vocab_df=vocab_df
    )
    
    # Create assay-type specific sheets
    if assay_type == 'metabarcoding':
        create_other_sheets(
            worksheets=worksheets,
            sheet_names=["experimentRunMetadata", "taxaRaw", "taxaFinal"],
            full_temp_file_name=full_temp_file_name,
            input_df=input_df,
            req_lev=req_lev,
            color_styles=color_styles,
            vocab_df=vocab_df
        )
    elif assay_type == 'targeted':
        create_other_sheets(
            worksheets=worksheets,
            sheet_names=["stdData", "eLowQuantData", "ampData"],
            full_temp_file_name=full_temp_file_name,
            input_df=input_df,
            req_lev=req_lev,
            color_styles=color_styles,
            vocab_df=vocab_df
        )

def create_readme_sheet(worksheet, input_file_name, req_lev, sample_type, assay_type,
                        project_id, assay_name, projectMetadata_user, sampleMetadata_user, color_styles, FAIRe_checklist_ver):
    """Create the README sheet with information about the template."""

    # Format ISO time like in R script
    now = datetime.now()
    iso_current_time = now.strftime("%Y-%m-%dT%H:%M:%S%z")
    if len(iso_current_time) >= 2:
        iso_current_time = re.sub(r"(\d{2})(\d{2})$", r"\1:\2", iso_current_time)

    # Build README content sections with new format (values below labels)
    readme1 = [
        ['FAIRe Checklist Version:'],
        [input_file_name],
        [''],
        ['Date/Time generated:'],
        [iso_current_time],
        ['']
    ]
    
    # Add Modification Timestamp section
    readme_timestamp_header = [
        ['Modification Timestamp:'],
        ['Sheet Name', 'Timestamp', 'Email']
    ]
    
    # Get all worksheet names except README and Drop-down values
    sheet_names = [ws.title for ws in worksheet.spreadsheet.worksheets() 
                  if ws.title not in ["README", "Drop-down values"]]
    
    # Create rows for each sheet (empty timestamp and email cells)
    readme_timestamp_rows = [[name, '', ''] for name in sheet_names]
    
    # Template parameters section
    readme2 = [
        [''],
        ['Template parameters:'],
        [f'project_id = {project_id}'],
        [f'assay_name = {" | ".join(assay_name)}'],
        [f'assay_type = {assay_type}'],
        [f'req_lev = {" | ".join(req_lev)}']
    ]
    
    if any(s.lower() == 'other' for s in sample_type):
        readme2.append(
            [f'sample_type = {" | ".join(sample_type)} '
             '(Note: this option provides sample-type-specific fields for ALL sample types)']
        )
    else:
        readme2.append(
            [f'sample_type = {" | ".join(sample_type)} '
             '(Note: this option provides sample-type-specific fields for the selected sample type(s))']
        )
    
    if projectMetadata_user:
        readme2.append([f'projectMetadata_user = {" | ".join(projectMetadata_user)}'])
    
    if sampleMetadata_user:
        readme2.append([f'sampleMetadata_user = {" | ".join(sampleMetadata_user)}'])
    
    readme2.append([''])
    
    # Requirement levels section
    readme3 = [
        ['Requirement levels:'],
        ['M = Mandatory'],
        ['HR = Highly recommended'],
        ['R = Recommended'],
        ['O = Optional'],
        ['']
    ]
    
    # Sheets in this Google sheet section (renamed from List of files)
    readme4 = [
        ['Sheets in this Google sheet:'],
        [f'projectMetadata_{project_id}'],
        [f'sampleMetadata_{project_id}']
    ]
    
    if assay_type == 'metabarcoding':
        readme4.extend([
            [f'experimentRunMetadata_{project_id}'],
            [f'otuRaw_{project_id}_{assay_name[0]}_<seq_run_id>' if len(assay_name) == 1 else f'otuRaw_{project_id}_<assay_name>_<seq_run_id>'],
            [f'otuFinal_{project_id}_{assay_name[0]}_<seq_run_id>' if len(assay_name) == 1 else f'otuFinal_{project_id}_<assay_name>_<seq_run_id>'],
            [f'taxaRaw_{project_id}_{assay_name[0]}_<seq_run_id>' if len(assay_name) == 1 else f'taxaRaw_{project_id}_<assay_name>_<seq_run_id>'],
            [f'taxaFinal_{project_id}_{assay_name[0]}_<seq_run_id>' if len(assay_name) == 1 else f'taxaFinal_{project_id}_<assay_name>_<seq_run_id>'],
            ['Note: otuRaw, otuFinal, taxaRaw and taxaFinal should be produced for each assay_name and seq_run_id'],
            ['Note: <seq_run_id> in the file names should match with seq_run_id in your experimentRunMetadata']
        ])
    elif assay_type == 'targeted':
        readme4.extend([
            [f'stdData_{project_id}'],
            [f'eLowQuantData_{project_id} (if applicable)'],
            [f'ampData_{project_id}_{assay_name[0]}' if len(assay_name) == 1 else f'ampData_{project_id}_<assay_name>'],
            ['Note: ampData should be produced for each assay_name']
        ])
    
    # Combine all readme sections
    readme_data = readme1 + readme_timestamp_header + readme_timestamp_rows + readme2 + readme3 + readme4
    
    # Write data to sheet - all at once to reduce API calls
    worksheet.update('A1', readme_data)
    
    # Format header rows (bold) using format_cell_ranges
    header_format = gsf.CellFormat(textFormat=gsf.TextFormat(bold=True))
    
    # Calculate row positions
    timestamp_section_start = 7
    timestamp_section_end = 8 + len(sheet_names)
    template_params_start = timestamp_section_end + 2
    req_levels_start = template_params_start + len(readme2) - 1
    sheets_section_start = req_levels_start + len(readme3)
    
    # Create a list of (range, format) tuples for section headers
    format_ranges = [
        ('A1:A1', header_format),  # FAIRe Checklist Version
        ('A4:A4', header_format),  # Date/Time generated
        ('A7:A7', header_format),  # Modification Timestamp
        ('A8:C8', header_format),  # Timestamp table headers
        (f'A{template_params_start}:A{template_params_start}', header_format),  # Template parameters
        (f'A{req_levels_start}:A{req_levels_start}', header_format),  # Requirement levels
        (f'A{sheets_section_start}:A{sheets_section_start}', header_format)   # Sheets in this Google sheet
    ]
    
    # Apply all formatting at once
    gsf.format_cell_ranges(worksheet, format_ranges)
    
    # Apply color formatting to requirement levels - separate call to ensure correct formatting
    req_level_rows = {
        'M': req_levels_start + 1,
        'HR': req_levels_start + 2,
        'R': req_levels_start + 3,
        'O': req_levels_start + 4
    }
    
    for level, row in req_level_rows.items():
        if level in color_styles:
            gsf.format_cell_range(worksheet, f'A{row}:A{row}', color_styles[level])

def create_dropdown_sheet(worksheet, vocab_df, assay_type, assay_name):
    """Create the Drop-down values sheet with controlled vocabulary options."""
    
    # Simply copy the Drop-down values sheet from the template
    if not vocab_df.empty:
        # Replace NaN values with empty strings to avoid JSON errors
        vocab_df_clean = vocab_df.fillna('')

        # Convert DataFrame to list of lists for gspread
        data = [vocab_df_clean.columns.tolist()] + vocab_df_clean.values.tolist()
        
        # Update the worksheet
        worksheet.update("A1", data)
        
        # Format header row
        header_format = gsf.CellFormat(textFormat=gsf.TextFormat(bold=True))
        gsf.format_cell_range(worksheet, "1:1", header_format)
        
        print("Created Drop-down values sheet")

def create_project_metadata_sheet(worksheet, full_temp_file_name, input_df, req_lev, assay_type,
                                 project_id, assay_name, projectMetadata_user, color_styles, vocab_df, FAIRe_checklist_ver):
    """Create and format the projectMetadata sheet."""
    
    # Read the projectMetadata sheet from the template
    project_meta_df = pd.read_excel(full_temp_file_name, sheet_name="projectMetadata")
    
    # Filter rows based on assay_type
    section2rm = []
    if assay_type == 'metabarcoding':
        section2rm = ['Targeted assay detection']
    elif assay_type == 'targeted':
        section2rm = ['Library preparation sequencing', 'Bioinformatics', 'OTU/ASV']
    
    for section in section2rm:
        project_meta_df = project_meta_df[project_meta_df['section'] != section]
    
    # Filter rows based on requirement levels
    req_lev2rm = [level for level in ['M', 'HR', 'R', 'O'] if level not in req_lev]
    for level in req_lev2rm:
        project_meta_df = project_meta_df[project_meta_df['requirement_level_code'] != level]
    
    # Add user-defined fields if provided
    if projectMetadata_user:
        user_rows = []
        for field in projectMetadata_user:
            user_row = {col: "" for col in project_meta_df.columns}
            user_row["term_name"] = field
            user_row["requirement_level_code"] = "O"  # Optional
            user_row["section"] = "User defined"
            user_rows.append(user_row)
        
        # Append user fields to the dataframe
        user_df = pd.DataFrame(user_rows)
        project_meta_df = pd.concat([project_meta_df, user_df], ignore_index=True)
    
    # Convert columns to string type to avoid dtype warnings
    project_meta_df['project_level'] = project_meta_df['project_level'].astype(str)
    project_meta_df.loc[project_meta_df['term_name'] == 'project_id', 'project_level'] = project_id

    # For assay columns
    for i, name in enumerate(assay_name):
        col_name = f"assay{i+1}"
        if col_name in project_meta_df.columns:
            project_meta_df[col_name] = project_meta_df[col_name].astype(str)
            project_meta_df.loc[project_meta_df['term_name'] == 'assay_name', col_name] = name
    
    # Replace NaN values with empty strings to avoid JSON errors
    project_meta_df_clean = project_meta_df.fillna('')

    # Convert DataFrame to list of lists for gspread
    data = [project_meta_df_clean.columns.tolist()] + project_meta_df_clean.values.tolist()

    # Resize worksheet to accommodate all data (add some buffer)
    rows_needed = len(data) + 10  # Add buffer
    cols_needed = len(data[0]) + 5  # Add buffer
    worksheet.resize(rows=rows_needed, cols=cols_needed)
    
    # Update the worksheet
    worksheet.update("A1", data)
    
    # Format headers and cells using format_cell_ranges
    header_format = gsf.CellFormat(textFormat=gsf.TextFormat(bold=True))
    
    # Create a list of (range, format) tuples
    format_ranges = [
        ('1:1', header_format)  # Format header row
    ]
    
    # Format term_name column with bold
    term_name_col = project_meta_df.columns.get_loc("term_name") + 1  # +1 for 1-indexing
    term_name_range = f"{chr(64 + term_name_col)}2:{chr(64 + term_name_col)}{len(data)}"
    format_ranges.append((term_name_range, header_format))
    
    # Format requirement level cells with colors
    req_level_col = project_meta_df.columns.get_loc("requirement_level_code") + 1  # +1 for 1-indexing
    
    # Instead of using DataFrame index, use row position in the data
    for row_idx in range(1, len(data)):  # Skip header row
        if row_idx < len(data) and req_level_col-1 < len(data[row_idx]):
            req_level = data[row_idx][req_level_col-1]  # Get value directly from data
            if req_level in color_styles:
                cell = f"{chr(64 + req_level_col)}{row_idx+1}"  # +1 for 1-indexing
                format_ranges.append((cell, color_styles[req_level]))
    
    # Apply all formatting at once
    gsf.format_cell_ranges(worksheet, format_ranges)
    
    # Batch all data validation requests
    validation_requests = []
    
    # Add dropdowns for controlled vocabulary fields
    for idx, row in project_meta_df.iterrows():
        term = row['term_name']
        # Find if this term has a dropdown
        vocab_row = vocab_df[vocab_df['term_name'] == term]
        if not vocab_row.empty:
            # Get the dropdown values
            n_options = int(vocab_row.iloc[0]['n_options'])
            values = [str(vocab_row.iloc[0][f'vocab{i+1}']) for i in range(n_options) if pd.notna(vocab_row.iloc[0][f'vocab{i+1}'])]
            
            if values:
                # Apply dropdown to project_level column
                project_level_col = project_meta_df.columns.get_loc("project_level") + 1
                cell = f"{chr(64 + project_level_col)}{idx+2}"
                
                # Create data validation rule using the correct API
                validation_rule = {
                    "setDataValidation": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": idx+1,  # 0-indexed in API
                            "endRowIndex": idx+2,
                            "startColumnIndex": project_level_col-1,  # 0-indexed in API
                            "endColumnIndex": project_level_col
                        },
                        "rule": {
                            "condition": {
                                "type": "ONE_OF_LIST",
                                "values": [{"userEnteredValue": v} for v in values]
                            },
                            "showCustomUi": True,
                            "strict": True
                        }
                    }
                }
                validation_requests.append(validation_rule)
                
                # If multiple assays, apply to each assay column
                if len(assay_name) > 1:
                    for i in range(len(assay_name)):
                        col_name = f"assay{i+1}"
                        if col_name in project_meta_df.columns:
                            assay_col = project_meta_df.columns.get_loc(col_name) + 1
                            
                            # Create data validation rule for assay column
                            assay_validation_rule = {
                                "setDataValidation": {
                                    "range": {
                                        "sheetId": worksheet.id,
                                        "startRowIndex": idx+1,  # 0-indexed in API
                                        "endRowIndex": idx+2,
                                        "startColumnIndex": assay_col-1,  # 0-indexed in API
                                        "endColumnIndex": assay_col
                                    },
                                    "rule": {
                                        "condition": {
                                            "type": "ONE_OF_LIST",
                                            "values": [{"userEnteredValue": v} for v in values]
                                        },
                                        "showCustomUi": True,
                                        "strict": True
                                    }
                                }
                            }
                            validation_requests.append(assay_validation_rule)
    
    # Apply all data validations in a single batch request
    if validation_requests:
        worksheet.spreadsheet.batch_update({'requests': validation_requests})
    
    # Batch all note requests
    note_requests = []
    
    # Add comments to cells - use the data directly instead of DataFrame index
    term_name_col = project_meta_df.columns.get_loc("term_name") + 1  # +1 for 1-indexing
    
    for row_idx in range(1, len(data)):  # Skip header row
        if row_idx < len(data) and term_name_col-1 < len(data[row_idx]):
            term = data[row_idx][term_name_col-1]  # Get term directly from data
            
            if term and term not in projectMetadata_user:
                # Find the term in the input dataframe
                term_rows = input_df[input_df['term_name'] == term]
                if not term_rows.empty:
                    term_info = term_rows.iloc[0]
                    
                    # Build comment text
                    comment_parts = []
                    
                    # Requirement level
                    req_level = term_info['requirement_level']
                    req_cond = term_info['requirement_level_condition']
                    if pd.isna(req_cond):
                        comment_parts.append(f"Requirement level: {req_level}")
                    else:
                        comment_parts.append(f"Requirement level: {req_level} ({req_cond})")
                    
                    # Description and example
                    comment_parts.append(f"Description: {term_info['description']}")
                    comment_parts.append(f"Example: {term_info['example']}")
                    
                    # Field type
                    field_type = term_info['term_type']
                    if field_type == 'controlled vocabulary':
                        vocab_options = term_info['controlled_vocabulary_options']
                        comment_parts.append(f"Field type: {field_type} ({vocab_options})")
                    elif field_type == 'fixed format':
                        format_spec = term_info['fixed_format']
                        comment_parts.append(f"Field type: {field_type} ({format_spec})")
                    else:
                        comment_parts.append(f"Field type: {field_type}")
                    
                    # Add the comment to the term_name cell
                    comment_text = "\n".join(comment_parts)
                    
                    # Create note request using the correct API
                    note_request = {
                        "updateCells": {
                            "range": {
                                "sheetId": worksheet.id,
                                "startRowIndex": row_idx,  # 0-indexed in API
                                "endRowIndex": row_idx+1,
                                "startColumnIndex": term_name_col-1,  # 0-indexed in API
                                "endColumnIndex": term_name_col
                            },
                            "rows": [{
                                "values": [{
                                    "note": comment_text
                                }]
                            }],
                            "fields": "note"
                        }
                    }
                    note_requests.append(note_request)
    
    # Apply all notes in a single batch request
    if note_requests:
        worksheet.spreadsheet.batch_update({'requests': note_requests})

def create_sample_metadata_sheet(worksheet, full_temp_file_name, input_df, req_lev, sample_type,
                               assay_type, assay_name, sampleMetadata_user, color_styles, vocab_df):
    """Create and format the sampleMetadata sheet."""
    
    # Use pandas DataFrame directly
    sheet_df = full_template_df["sampleMetadata"]
    # Replace NaN values with empty strings to avoid JSON errors
    sheet_df = sheet_df.fillna('')
    data = sheet_df.values.tolist()
    
    # Find the term_name row
    term_name_row_idx = None
    for i, row in enumerate(data):
        if row and 'samp_name' in row:
            term_name_row_idx = i
            break
    
    if term_name_row_idx is None:
        raise ValueError("Could not find samp_name in sampleMetadata template")
    
    # Get the column headers (term names)
    headers = data[term_name_row_idx]
    
    # Filter columns based on sample_type
    if not any(s.lower() == 'other' for s in sample_type):
        # Find sample_type specific columns to keep
        cols_to_keep = []
        for i, col in enumerate(headers):
            if col:
                # Find this column in the input dataframe
                col_info = input_df[input_df['term_name'] == col]
                if not col_info.empty:
                    sample_type_spec = col_info.iloc[0]['sample_type_specificity']
                    if pd.isna(sample_type_spec) or sample_type_spec == 'ALL' or any(s in str(sample_type_spec).split(',') for s in sample_type):
                        cols_to_keep.append(i)
                else:
                    # Keep columns not in the input dataframe (like user-defined)
                    cols_to_keep.append(i)
        
        # Create a filtered data list with only the columns to keep
        filtered_data = []
        for row in data:
            filtered_row = [row[i] for i in cols_to_keep]
            filtered_data.append(filtered_row)
        data = filtered_data
    
    # Find requirement level row
    req_level_row_idx = None
    for i, row in enumerate(data):
        if row and '# requirement_level_code' in row:
            req_level_row_idx = i
            break
    
    # Filter by requirement level
    if req_level_row_idx is not None:
        req_levels = data[req_level_row_idx]
        cols_to_keep = []
        for i, level in enumerate(req_levels):
            if level in req_lev or not level:
                cols_to_keep.append(i)
        
        # Create a filtered data list with only the columns to keep
        filtered_data = []
        for row in data:
            filtered_row = [row[i] for i in cols_to_keep]
            filtered_data.append(filtered_row)
        data = filtered_data
    
    # Handle detected_notDetected for targeted assays with multiple assay names
    if assay_type == 'targeted' and len(assay_name) > 1:
        # Find the detected_notDetected column
        detected_col_idx = None
        for i, col in enumerate(data[term_name_row_idx]):
            if col == 'detected_notDetected':
                detected_col_idx = i
                break
        
        if detected_col_idx is not None:
            # Remove the original column
            for i in range(len(data)):
                data[i].pop(detected_col_idx)
            
            # Add new columns for each assay
            for name in assay_name:
                for i in range(len(data)):
                    if i == term_name_row_idx:
                        data[i].append(f'detected_notDetected_{name}')
                    elif i == req_level_row_idx:
                        # Copy the requirement level from the original column
                        data[i].append(req_levels[detected_col_idx] if detected_col_idx < len(req_levels) else '')
                    else:
                        data[i].append('')
    
    # Add user-defined fields if provided
    if sampleMetadata_user:
        section_row_idx = None
        for i, row in enumerate(data):
            if row and '# section' in row:
                section_row_idx = i
                break
        
        for field in sampleMetadata_user:
            for i in range(len(data)):
                if i == term_name_row_idx:
                    data[i].append(field)
                elif i == req_level_row_idx:
                    data[i].append('O')  # Optional for user fields
                elif i == section_row_idx:
                    data[i].append('User defined')
                else:
                    data[i].append('')
    
    # Resize worksheet to accommodate all data (add some buffer)
    rows_needed = len(data) + 20  # Add buffer
    cols_needed = len(data[0]) + 10 if data else 50  # Add buffer
    worksheet.resize(rows=rows_needed, cols=cols_needed)
    
    # Write data to Google Sheets
    worksheet.update("A1", data)
    
    # Format headers - make the term_name row bold
    term_name_range = f"{term_name_row_idx+1}:{term_name_row_idx+1}"
    header_format = gsf.CellFormat(textFormat=gsf.TextFormat(bold=True))
    gsf.format_cell_range(worksheet, term_name_range, header_format)
    
    # Format requirement level cells with colors
    req_level_col = req_level_row_idx + 1
    for i, level in enumerate(req_levels):
        if level in color_styles:
            cell = f"{chr(64 + req_level_col)}{i+2}"
            gsf.format_cell_range(worksheet, cell, color_styles[level])

def create_other_sheets(worksheets, sheet_names, full_temp_file_name, input_df, req_lev, color_styles, vocab_df):
    """Create and format other sheets based on assay type."""
    
    for sheet_name in sheet_names:
        worksheet = worksheets[sheet_name]
        
        # Read the template using pandas instead of openpyxl
        sheet_df = pd.read_excel(full_temp_file_name, sheet_name=sheet_name, header=None)
        
        # Replace NaN values with empty strings to avoid JSON errors
        sheet_df = sheet_df.fillna('')
        
        # Convert to list of lists for gspread
        data = sheet_df.values.tolist()
        
        # Resize worksheet to accommodate all data (add some buffer)
        rows_needed = len(data) + 20  # Add buffer
        cols_needed = len(data[0]) + 10 if data else 50  # Add buffer
        worksheet.resize(rows=rows_needed, cols=cols_needed)
        
        # Write to Google Sheets
        worksheet.update("A1", data)
