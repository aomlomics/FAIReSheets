"""
Module for creating the README sheet in FAIReSheets.
"""

import re
from datetime import datetime
import gspread_formatting as gsf

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
             '(Note: this option provides sample-type-specific fields for the selected sample type(s))']
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