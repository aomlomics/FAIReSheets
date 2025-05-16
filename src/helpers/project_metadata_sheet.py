"""
Module for creating the projectMetadata sheet in FAIReSheets.
"""

import pandas as pd
import gspread_formatting as gsf

def create_project_metadata_sheet(worksheet, full_temp_file_name, input_df, req_lev, assay_type,
                                  project_id, assay_name, projectMetadata_user, color_styles, vocab_df, FAIRe_checklist_ver):
    """Create and format the projectMetadata sheet."""
    
    # Read the projectMetadata sheet from the template
    project_meta_df = pd.read_excel(full_temp_file_name, sheet_name="projectMetadata")
    
    # Replace NaN values with empty strings immediately after loading
    project_meta_df = project_meta_df.fillna('')
    
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
    for col in project_meta_df.columns:
        project_meta_df[col] = project_meta_df[col].astype(str)
    
    # Replace 'nan' strings with empty strings
    project_meta_df = project_meta_df.replace('nan', '')
    
    # Get the input file name for the checklist version
    input_file_name = f'FAIRe_checklist_{FAIRe_checklist_ver}.xlsx'
    
    # Pre-fill values from config.yaml
    project_meta_df.loc[project_meta_df['term_name'] == 'project_id', 'project_level'] = project_id
    project_meta_df.loc[project_meta_df['term_name'] == 'assay_type', 'project_level'] = assay_type
    project_meta_df.loc[project_meta_df['term_name'] == 'checkls_ver', 'project_level'] = input_file_name
    
    # Handle assay_name based on number of assays
    if len(assay_name) == 1:
        # Single assay - just put it in project_level
        project_meta_df.loc[project_meta_df['term_name'] == 'assay_name', 'project_level'] = assay_name[0]
    else:
        # Multiple assays - put pipe-separated list in project_level
        project_meta_df.loc[project_meta_df['term_name'] == 'assay_name', 'project_level'] = ' | '.join(assay_name)
        
        # Also create individual columns for each assay
        for i, name in enumerate(assay_name):
            # Use the assay name as the column name instead of "assay{i+1}"
            col_name = name
            if col_name not in project_meta_df.columns:
                project_meta_df[col_name] = ""
            project_meta_df.loc[project_meta_df['term_name'] == 'assay_name', col_name] = name
    
    # Final check to ensure no 'nan' values remain
    project_meta_df = project_meta_df.replace('nan', '')
    
    # Convert DataFrame to list of lists for gspread
    data = [project_meta_df.columns.tolist()] + project_meta_df.values.tolist()
    
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