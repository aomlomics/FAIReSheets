<div align="center">
  <img src="src/helpers/banner_fairesheets.png" alt="FAIReSheets Banner" width="800">
</div>

FAIReSheets converts the FAIR eDNA ([FAIRe](https://fair-edna.github.io/index.html)) data checklist to customizable Google Sheets templates. FAIReSheets can be run in one of 2 modes:
1. **FAIRe eDNA:** Generate FAIRe eDNA data templates from the FAIRe checklist
2. **ODE-ready (NOAA):** Generate (FAIRe-NOAA) data templates, formatted for submission to the [Ocean DNA Explorer](https://www.oceandnaexplorer.org/), and can be used as input to [edna2obis](https://github.com/aomlomics/edna2obis), a data pipeline for submission to [GBIF](https://www.gbif.org/) and [OBIS](https://obis.org/).

NOTE: FAIReSheets generates BLANK templates. You must fill them in with data manually after they're generated.

### Quick Start Summary
Email bayden.willms@noaa.gov to be added to the user list and receive the link to the credentials file, create a blank Google Sheet, configure the `.env` file with the Google Sheet ID and the Git Gist URL, specify your parameters in `config.yaml` and optionally in `NOAA_config.yaml` if you want NOAA/ODE-ready templates, follow the authentication workflow (on your browser), and run `python run.py`.

---
### Table of Contents
1. [Prerequisites](#Prerequisites)
2. [Installation](#Installation)
3. [Configuration](#Configuration)
4. [Usage](#Usage)
5. [Troubleshooting](#Troubleshooting)
---

### Prerequisites
Before using FAIReSheets, you'll need to request access. This only needs to happen once:
1. **Request Access**: 
   - Email bayden.willms@noaa.gov with the subject "FAIReSheets Access Request".
   - Include the email address associated with your Google account. For NOAA users, use your @noaa.gov email.
2. **Receive Credentials**: 
   - Once approved, you'll receive an email with a link to a private Git Gist. This Gist contains the `client_secrets.json` and `token.json`, which are files needed for authentication.

### Installation
1. **Clone the Repository**:
   ```bash
   git clone https://github.com/aomlomics/FAIReSheets.git
   cd FAIReSheets
   ```
2. **Set up the Environment**:
   - Install Conda if you don't have it already.
   - Create and activate the Conda environment:
     ```bash
     conda env create -f environment.yml
     conda activate FAIRe
     ```

### Configuration
1. **Create `.env` file**:
   - In the FAIReSheets directory, create a `.env` file. Or alternatively, if you run FAIReSheets without having a `.env` file, one will be created for you. Note that you will still need to fill in the Git Gist URL and Spreadsheet ID to that `.env`.
2. **Configure `.env` file**:
   - Open the `.env` file and add the following, replacing the placeholder text with your actual information:
     ```
     SPREADSHEET_ID=your_spreadsheet_id_here
     GIST_URL=your_gist_url_here
     ```
   - `SPREADSHEET_ID`: This is the ID of the Google Sheet you want to populate. You can find it in the URL of your Google Sheet, between the **/d/** and **/edit**: `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit`.
   - `GIST_URL`: The GIST_URL will be sent to you via email after you've been granted access to FAIReSheets (see first section).

3. **Customize your FAIRe checklist:**
   - The FAIRe data checklist is designed to be customizable. If you have data fields that are not included in the checklist, you can manually add them into the checklist as User Defined fields, and your changes will be reflected in the templates you generate. We recommend trying your best to align your custom fields with fields in existing eDNA data standards, like Darwin Core or MIXs.

### Usage

FAIReSheets can generate **EITHER** FAIRe eDNA templates, **OR** ODE-ready (NOAA) eDNA templates. See more info in the bullet points below:
- **FAIRe eDNA:** The default mode of FAIReSheets, this will generate the exact format which the [FAIRe eDNA collaboration](https://fair-edna.github.io/index.html) supports. Parameters for the template are set in the `config.yaml` file. HINT: for users with qPCR data, this is what you want!
- **(NOAA) ODE-ready:** Generates templates in the format required for submission to the [Ocean DNA Explorer](https://www.oceandnaexplorer.org/), NOAA 'Omics own eDNA data portal. Submission to ODE will unlock your data's potential, with an intuitive user interface, data visualizations, search, API endpoints, and more! These data templates can also be used as input to [edna2obis](https://github.com/aomlomics/edna2obis), a data pipeline for submission to [GBIF](https://www.gbif.org/) and [OBIS](https://obis.org/). This workflow now also generates a `taxaFinal` sheet. Parameters for the template are set in the `NOAA_config.yaml` file. HIGHLY RECOMMENDED to any user with metabarcoding / targeted eDNA data!


Customize your generated data templates depending on your data:
 - Open `config.yaml` and `NOAA_config.yaml` to set your project-specific parameters.
 - Comments in the files explain what each parameter does.
 - `config.yaml` is for FAIRe eDNA template parameters.
 - `NOAA_config.yaml` is for NOAA / ODE-ready template parameters.

If you would like to generate NOAA / ODE-ready templates, set the `run_noaa_formatting` config parameter to `true`, like this: 
```bash 
run_noaa_formatting: true
```

You are now ready to run the code! Run FAIReSheets from the root project directory using: 
```bash
python run.py
```
This will:
1. Generate the data templates in your specified Google Sheet.
2. If `run_noaa_formatting` is `true` in `NOAA_config.yaml`, it will then format the sheet for the Ocean DNA Explorer.

#### First-Time Authentication
When you run FAIReSheets for the first time, the following will happen:
1. A browser window will open, prompting you to log in to your Google account. 
   - **Important**: Use the same Google account you provided when requesting access.
2. After logging in, you'll be asked to grant permission to FAIReSheets to access your Google Sheets
3. Once you grant permission, a `token.json` file will be created in the project directory. This file stores your authentication token, so you won't have to log in every time you run the tool.

### Troubleshooting
- **Problem**: "User is not on the approved list"
  - **Solution**: You need to be on the approved users list to run this tool. Email bayden.willms@noaa.gov to request access.
- **Problem**: Authentication errors (e.g., "invalid_grant")
  - **Solution**: Delete the `token.json` file and run the tool again. This will re-trigger the authentication process.
  - **Solution**: Make sure you're using the Google account email you provided when requesting access, and that you checked the boxes to allow FAIReSheets to edit Google Sheets in your Google Drive.
- **Problem**: Errors when running FAIReSheets
  - **Solution**: Make sure the Google Sheet that FAIReSheets is editing is **EMPTY**. You can use Google Drive's built in Restore History button before running FAIReSheets again, or, make a new Google Sheet and replace the Spreadsheet ID in the `.env` file. 
- **Problem**: Missing `client_secrets.json` or `token.json`
  - **Solution**: Make sure you've downloaded these files from the Git Gist and placed them in the root of the project directory.

## Optional (recommended): Google Apps Script

Copy and Paste the following Google Apps Script for some helpful features, like tracking modifications in the README of your Google Sheet, data validation on important fields, and a button to download all sheets as TSV files (needed for [Ocean DNA Explorer](https://www.oceandnaexplorer.org/) and edna2obis submission).

NOTE: FAIReSheets now standardizes font family + font size across all sheets during template generation, so the font-related Apps Script features are optional.

### Adding the Google Apps Script

1. Open your Google Sheet.
2. Click on `Extensions` in the menu, then select `Apps Script`.
3. Delete any code in the script editor and copy-paste the following code.
   IMPORTANT: Make sure you copy the FULL script starting from `function onOpen()` (the menu will NOT appear if you only paste `exportSheetsAsTsv()`).
   After saving, reload/refresh the Google Sheet tab to trigger `onOpen()` and show a new `FAIReSheets Tools` menu in the Google Sheets UI.

Once the `FAIReSheets Tools` menu appears, you can:
- Download all sheets as TSV files (for ODE / edna2obis submission)
- Standardize font across all sheets
- Reorder `projectMetadata`, `sampleMetadata`, `experimentRunMetadata`, and all `analysisMetadata*` sheets (by moving rows/columns in-place)

The reordering tool:
- Can be run **before or after** you’ve filled the sheet with data (it moves entire rows/columns, so your entered data moves with the fields)
- Uses clean ordered lists inside the Apps Script (you can edit them if you want)
- Will **not** break if you include terms/headers that aren’t present in your sheet (it will skip them)
- Reports what it moved and what was missing (ignored)
   The first time you run it, Google may ask you to authorize permissions.

```javascript
/**
 * Adds a custom menu to the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('FAIReSheets Tools')
      .addItem('Download all sheets as TSV', 'exportSheetsAsTsv')
      .addItem('Standardize font across all sheets', 'standardizeFontAcrossAllSheets')
      .addSeparator()
      .addItem('Reorder metadata sheets (column/field order)', 'reorderMetadataSheets')
      .addToUi();

  // Auto-run font standardization once (no user action required).
  // Uses Document Properties so it runs only the first time after the script is added/updated.
  standardizeFontAcrossAllSheetsOnce_();
}

/**
 * Column/field ordering lists (optional).
 *
 * - sampleMetadata / experimentRunMetadata: these are COLUMN headers to move to the front in this order.
 * - projectMetadata / analysisMetadata: these are term_name values (ROWS) to move to the top in this order.
 *
 * Anything not listed is left in place (and the script will NOT error if a listed item is missing).
 *
 * Tip: you can keep these short (just "pin" your most important fields first), or make them exhaustive.
 */
const COLUMN_OR_FIELD_ORDER = {
  projectMetadata: [
// Core project information
    "project_id",
    "project_name",
    "parent_project_id",
    "institution",
    "institutionID",
    "recordedBy",
    "recordedByID",
    "project_contact",
    "sample_type",    
    "study_factor",
    "expedition_id",
    "ship_crs_expocode",
    "woce_sect",
    "bioproject_accession",
    "projectDescription", 
    "dataDescription",
// Data Management Info    
    "checkls_ver",
    "mod_date",
    "license",
    "rightsHolder",
    "accessRights",
    "informationWithheld",
    "dataGeneralizations",
    "bibliographicCitation",
    "associated_resource",
    "code_repo",
    "biological_rep",
// Assay and PCR Info
    "assay_name",
    "assay_type",
    "assay_name_alternate",
    "assay_reference",
    "sterilise_method",
    "neg_cont_0_1",
    "pos_cont_0_1",
    "pcr_primer_forward",
    "pcr_primer_reverse",
    "pcr_primer_name_forward",
    "pcr_primer_name_reverse",
    "pcr_primer_reference_forward",
    "pcr_primer_reference_reverse",
    "pcr_primer_name_published_forward",
    "pcr_primer_name_published_reverse",
    "pcr_primer_vol_forward",
    "pcr_primer_vol_reverse",
    "pcr_primer_conc_forward",
    "pcr_primer_conc_reverse",
    "pcr_0_1",
    "inhibition_check_0_1",
    "inhibition_check",
    "targetTaxonomicAssay",
    "targetTaxonomicScope",
    "target_gene",
    "target",
  ],
  sampleMetadata: [
    "samp_name",
    "sample_type",
    "samp_category",
    "pos_cont_type",
    "neg_cont_type",
    "short_name",
    "expedition_id",
    "expeditionName",
    "expeditionURL",
    "expeditionStartDate",
    "expeditionEndDate",
    "expeditionLocation",
    "materialSampleID",
    "sample_derived_from",
    "sample_composed_of",
    "rel_cont_id",
    "biological_rep_relation",
    "eventDate",
    "eventDurationValue",
    "verbatimEventDate",
    "verbatimEventTime",
    "country",
    "geo_loc_name",
    "locality",
    "env_medium",
    "env_broad_scale",
    "env_local_scale",
  ],
  experimentRunMetadata: [
    "samp_name",
    "lib_id",
    "assay_name",
    "seq_run_id",
    "pcr_plate_id",
  ],
  analysisMetadata: [
    "project_id",
    "assay_name",
    "analysis_run_name",
  ],
};

/**
 * Runs standardizeFontAcrossAllSheets() only once per spreadsheet (unless you change the key).
 */
function standardizeFontAcrossAllSheetsOnce_() {
  const props = PropertiesService.getDocumentProperties();
  const key = "FAIRESHEETS_FONT_STANDARDIZED_V1";
  if (props.getProperty(key) === "true") return;

  standardizeFontAcrossAllSheets();
  props.setProperty(key, "true");
}

/**
 * Standardizes font for every cell in every sheet.
 * This intentionally changes ONLY font family + font size (not background, bold, etc.).
 */
function standardizeFontAcrossAllSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  // Google Sheets default formatting is typically Arial 10.
  // If you want a different font, change "Arial" here.
  const fontFamily = "Arial";
  const fontSize = 10;

  sheets.forEach(sheet => {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    if (maxRows < 1 || maxCols < 1) return;

    sheet
      .getRange(1, 1, maxRows, maxCols)
      .setFontFamily(fontFamily)
      .setFontSize(fontSize);
  });
}

/**
 * Exports all sheets in the spreadsheet as TSV files to an auto-created Google Drive folder.
 */
function exportSheetsAsTsv() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetName = spreadsheet.getName();
  
  // Create folder name with timestamp
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const folderName = spreadsheetName + "_TSVs_" + timestamp;
  
  // Show confirmation dialog
  const confirmMessage = 'This will create a new folder in your Google Drive home directory named:\n\n"' + folderName + '"\n\nAll sheets will be exported as TSV files to this folder.\n\nContinue?';
  const response = ui.alert('Export Sheets as TSV', confirmMessage, ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    return; // User cancelled
  }
  
  // Create new folder in Drive root (each export gets its own timestamped folder)
  const rootFolder = DriveApp.getRootFolder();
  const folder = rootFolder.createFolder(folderName);
  
  const sheets = spreadsheet.getSheets();
  const filesCreated = [];
  const errors = [];

  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    const fileName = `${spreadsheetName}_${sheetName}.tsv`;
    const data = sheet.getDataRange().getValues();
    const tsvContent = data.map(row => 
      row.map(cell => cell.toString().replace(/[\t\n]/g, ' ')).join('\t')
    ).join('\n');

    try {
      let existingFiles = folder.getFilesByName(fileName);
      if (existingFiles.hasNext()) {
        existingFiles.next().setContent(tsvContent);
      } else {
        folder.createFile(fileName, tsvContent, MimeType.PLAIN_TEXT);
      }
      filesCreated.push(fileName);
    } catch (e) {
      errors.push(`Error for sheet "${sheetName}": ${e.message}`);
    }
  }

  let message = '';
  if (filesCreated.length > 0) {
    message += `Successfully exported ${filesCreated.length} sheets to folder "${folderName}" in your Google Drive.\n\nFiles:\n${filesCreated.join('\n')}`;
  }
  if (errors.length > 0) {
    message += `\n\nErrors encountered:\n${errors.join('\n')}`;
  }
  
  ui.alert(message || 'No sheets were exported.');
}

/**
 * Runs automatically when a user edits the spreadsheet.
 * Handles timestamp updates, data validation, and font standardization.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const spreadsheet = sheet.getParent();
  
  // Sheets to ignore for all operations
  const excludedSheets = ["README", "Drop-down values"];
  if (excludedSheets.includes(sheetName)) {
    return;
  }

  // 1. Update Modification Timestamps
  updateModificationTimestamp(spreadsheet, sheetName);
  
  // 2. Standardize Font
  range.setFontFamily("Arial").setFontSize(10);

  // 3. Perform Data Validation
  validateSheetData(spreadsheet);
}

/**
 * Updates the 'Last Modified' and 'Modified By' fields in the README sheet.
 */
function updateModificationTimestamp(spreadsheet, editedSheetName) {
  const readmeSheet = spreadsheet.getSheetByName("README");
  if (!readmeSheet) return;

  const data = readmeSheet.getDataRange().getValues();
  let timestampRowStart = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "Sheets in this Google Sheet:") {
      timestampRowStart = i + 2; // Start of sheet list, skipping header
      break;
    }
  }

  if (timestampRowStart === -1) return;

  for (let i = timestampRowStart; i < data.length; i++) {
    if (data[i][0] === editedSheetName) {
      const sheetRow = i + 1;
      const timestamp = new Date().toISOString();
      readmeSheet.getRange(sheetRow, 2).setValue(timestamp);
      readmeSheet.getRange(sheetRow, 3).setValue(Session.getActiveUser().getEmail());
      break;
    }
  }
}

/**
 * Performs several data validation checks across metadata sheets.
 */
function validateSheetData(spreadsheet) {
  const projectMetadataSheet = spreadsheet.getSheetByName("projectMetadata");
  const analysisMetadataSheets = spreadsheet.getSheets().filter(s => s.getName().startsWith("analysisMetadata_"));

  if (!projectMetadataSheet || analysisMetadataSheets.length === 0) {
    return; // Exit if required sheets don't exist
  }

  // Get project_id from projectMetadata
  const projectIdCell = findCellByValue(projectMetadataSheet, "project_id");
  if (!projectIdCell) return;
  const projectId = projectMetadataSheet.getRange(projectIdCell.row, projectIdCell.col + 1).getValue();

  // Get assay_name values from projectMetadata
  const assayNameCell = findCellByValue(projectMetadataSheet, "assay_name");
  if (!assayNameCell) return;
  const assayNames = projectMetadataSheet.getRange(assayNameCell.row, assayNameCell.col + 1).getValue().toString().split("|").map(name => name.trim());
  
  clearErrorFormatting(spreadsheet); // Clear previous errors before re-validating

  let foundAssayNames = new Set();
  
  // Loop through analysis sheets once to perform all checks
  analysisMetadataSheets.forEach(analysisSheet => {
    // Check 1: project_id must match projectMetadata
    const analysisProjectIdCell = findCellByValue(analysisSheet, "project_id");
    if (analysisProjectIdCell) {
      const analysisProjectId = analysisSheet.getRange(analysisProjectIdCell.row, analysisProjectIdCell.col + 1).getValue();
      if (analysisProjectId !== projectId) {
        addErrorFormatting(analysisSheet, analysisProjectIdCell.row, analysisProjectIdCell.col + 1, "Project ID must match the one in projectMetadata sheet");
      }
    }

    // Check 2: assay_name must be one of the values from projectMetadata
    const analysisAssayNameCell = findCellByValue(analysisSheet, "assay_name");
    if (analysisAssayNameCell) {
      const analysisAssayName = analysisSheet.getRange(analysisAssayNameCell.row, analysisAssayNameCell.col + 1).getValue();
      if (!assayNames.includes(analysisAssayName)) {
        addErrorFormatting(analysisSheet, analysisAssayNameCell.row, analysisAssayNameCell.col + 1, "Assay name must match one of the values in projectMetadata sheet");
      } else {
        foundAssayNames.add(analysisAssayName);
      }
    }

    // Check 3: analysis_run_name must be unique within the sheet
    const analysisRunNameCell = findCellByValue(analysisSheet, "analysis_run_name");
    if (analysisRunNameCell) {
      const data = analysisSheet.getDataRange().getValues();
      const runNameColIndex = analysisRunNameCell.col - 1;
      const analysisRunNames = new Set();
      for (let i = 0; i < data.length; i++) {
        const value = data[i][runNameColIndex];
        if (value && value.toString().trim() !== "") {
          if (analysisRunNames.has(value)) {
            addErrorFormatting(analysisSheet, i + 1, analysisRunNameCell.col, "Duplicate analysis_run_name found in this sheet");
          } else {
            analysisRunNames.add(value);
          }
        }
      }
    }
  });

  // Check 4: All assay_names from projectMetadata must be used
  assayNames.forEach(assayName => {
    if (assayName && !foundAssayNames.has(assayName)) {
      addErrorFormatting(projectMetadataSheet, assayNameCell.row, assayNameCell.col + 1, `Assay name "${assayName}" must be used in an analysisMetadata sheet`);
    }
  });
}

/**
 * Finds the first cell in a sheet that matches a given value.
 * @returns {{row: number, col: number}|null} Cell coordinates or null if not found.
 */
function findCellByValue(sheet, searchValue) {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].length; j++) {
      if (data[i][j] === searchValue) {
        return {row: i + 1, col: j + 1};
      }
    }
  }
  return null;
}

/**
 * Adds red background and an error note to a specified cell.
 */
function addErrorFormatting(sheet, row, col, message) {
  const cell = sheet.getRange(row, col);
  const currentNote = cell.getNote();
  const errorMessage = "ERROR: " + message;
  
  if (!currentNote.includes(errorMessage)) {
    cell.setNote(currentNote + (currentNote ? "\n" : "") + errorMessage);
  }
  cell.setBackground("#ff0000"); // Red background
}

/**
 * Clears all error-related formatting (red backgrounds and error notes) from all sheets.
 */
function clearErrorFormatting(spreadsheet) {
  const sheets = spreadsheet.getSheets();
  sheets.forEach(sheet => {
    const dataRange = sheet.getDataRange();
    const backgrounds = dataRange.getBackgrounds();
    const notes = dataRange.getNotes();
    
    for (let i = 0; i < backgrounds.length; i++) {
      for (let j = 0; j < backgrounds[i].length; j++) {
        if (backgrounds[i][j] === "#ff0000") { // If it's our error background color
          const cell = sheet.getRange(i + 1, j + 1);
          const note = notes[i][j];
          if (note) {
            const nonErrorLines = note.split("\n").filter(line => !line.startsWith("ERROR:"));
            cell.setNote(nonErrorLines.join("\n"));
          }
          cell.setBackground(null); // Remove background color
        }
      }
    }
  });
}

/**
 * Reorders only the main metadata sheets by moving entire columns/rows in-place,
 * so data validation, conditional formatting, notes, etc. move with them.
 */
function reorderMetadataSheets() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const confirmMessage =
    "This will reorder columns/fields in:\n" +
    "- projectMetadata\n" +
    "- sampleMetadata\n" +
    "- experimentRunMetadata\n" +
    "- analysisMetadata* (all sheets whose name starts with 'analysisMetadata')\n\n" +
    "It moves existing rows/columns in-place, skips missing fields, and preserves dropdowns and formatting.\n\n" +
    "Continue?";

  const response = ui.alert("Reorder metadata sheets", confirmMessage, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const results = [];

  // projectMetadata: move term_name rows
  {
    const sheet = spreadsheet.getSheetByName("projectMetadata");
    if (sheet) {
      results.push(reorderLongFormByTermName_(sheet, COLUMN_OR_FIELD_ORDER.projectMetadata, "projectMetadata"));
    } else {
      results.push('Skipped "projectMetadata" (sheet not found).');
    }
  }

  // sampleMetadata: move columns by header row
  {
    const sheet = spreadsheet.getSheetByName("sampleMetadata");
    if (sheet) {
      results.push(reorderWideFormByHeader_(sheet, COLUMN_OR_FIELD_ORDER.sampleMetadata, "sampleMetadata"));
    } else {
      results.push('Skipped "sampleMetadata" (sheet not found).');
    }
  }

  // experimentRunMetadata: move columns by header row
  {
    const sheet = spreadsheet.getSheetByName("experimentRunMetadata");
    if (sheet) {
      results.push(reorderWideFormByHeader_(sheet, COLUMN_OR_FIELD_ORDER.experimentRunMetadata, "experimentRunMetadata"));
    } else {
      results.push('Skipped "experimentRunMetadata" (sheet not found).');
    }
  }

  // analysisMetadata*: move term_name rows for all matching sheets
  {
    const analysisSheets = spreadsheet.getSheets().filter(s => s.getName().startsWith("analysisMetadata"));
    if (analysisSheets.length === 0) {
      results.push('Skipped "analysisMetadata*" (no matching sheets found).');
    } else {
      analysisSheets.forEach(s => {
        results.push(reorderLongFormByTermName_(s, COLUMN_OR_FIELD_ORDER.analysisMetadata, s.getName()));
      });
    }
  }

  ui.alert(results.filter(Boolean).join("\n\n"));
}

function isWideMetadataSheetLayout_(sheet) {
  // sampleMetadata / experimentRunMetadata layout:
  // row 1 col 1: "# requirement_level_code"
  // row 2 col 1: "# section"
  // row 3: term/field headers (e.g., samp_name in column 1)
  const a1 = (sheet.getRange(1, 1).getValue() || "").toString().trim();
  const a2 = (sheet.getRange(2, 1).getValue() || "").toString().trim();
  return a1 === "# requirement_level_code" && a2 === "# section";
}

function reorderWideFormByHeader_(sheet, desiredHeaderOrder, labelForMessages) {
  if (!desiredHeaderOrder || desiredHeaderOrder.length === 0) {
    return `No column order list provided for "${labelForMessages}". Nothing changed.`;
  }

  const lastCol = sheet.getLastColumn();
  const maxRows = sheet.getMaxRows();
  if (lastCol < 1 || maxRows < 1) return `Skipped "${labelForMessages}" (empty sheet).`;

  const isWideMeta = isWideMetadataSheetLayout_(sheet);
  const headerRow = isWideMeta ? 3 : findBestHeaderRow_(sheet, desiredHeaderOrder, 50);
  const headerValues = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(v => (v || "").toString().trim());

  const colByHeader = {};
  headerValues.forEach((h, idx) => {
    if (!h) return;
    if (colByHeader[h] == null) colByHeader[h] = idx + 1; // 1-based
  });

  const moved = [];
  const missing = [];

  // In wide-metadata layout, column 1 is special (it contains the row labels in rows 1-2,
  // and the leftmost key field in row 3). We keep it fixed so the sheet doesn't "break".
  let destCol = isWideMeta ? 2 : 1;
  desiredHeaderOrder.forEach(header => {
    const key = (header || "").toString().trim();
    if (!key) return;

    const srcCol = colByHeader[key];
    if (srcCol == null) {
      missing.push(key);
      return;
    }

    // Never move the first column on wide-metadata sheets.
    // If the user includes the column-1 header (e.g., samp_name) in the list, treat it as already placed.
    if (isWideMeta && srcCol === 1) {
      moved.push(key);
      return;
    }

    if (srcCol !== destCol) {
      sheet.moveColumns(sheet.getRange(1, srcCol, maxRows, 1), destCol);
      updateIndexMapAfterMove_(colByHeader, srcCol, destCol);
    }
    moved.push(key);
    destCol += 1;
  });

  return [
    `Reordered "${labelForMessages}" using header row ${headerRow}.`,
    moved.length ? `Moved to front (in order): ${moved.join(", ")}` : "No columns moved.",
    missing.length ? `Missing (ignored): ${missing.join(", ")}` : "",
  ].filter(Boolean).join("\n");
}

function reorderLongFormByTermName_(sheet, desiredTermOrder, labelForMessages) {
  if (!desiredTermOrder || desiredTermOrder.length === 0) {
    return `No field order list provided for "${labelForMessages}". Nothing changed.`;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const maxCols = sheet.getMaxColumns();
  if (lastRow < 2 || lastCol < 1) return `Skipped "${labelForMessages}" (no data rows).`;

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => (v || "").toString().trim());
  let termNameCol = header.indexOf("term_name") + 1; // 1-based
  if (termNameCol < 1) {
    // Fall back: search a bit for "term_name" if headers aren't on row 1 for some reason.
    termNameCol = findCellByValue(sheet, "term_name")?.col || 0;
  }
  if (termNameCol < 1) return `Skipped "${labelForMessages}" (could not find "term_name" column).`;

  const termValues = sheet.getRange(2, termNameCol, lastRow - 1, 1).getValues().map(r => (r[0] || "").toString().trim());
  const rowByTerm = {};
  termValues.forEach((t, i) => {
    if (!t) return;
    if (rowByTerm[t] == null) rowByTerm[t] = i + 2; // actual row number
  });

  const moved = [];
  const missing = [];

  let destRow = 2;
  desiredTermOrder.forEach(term => {
    const key = (term || "").toString().trim();
    if (!key) return;

    const srcRow = rowByTerm[key];
    if (srcRow == null) {
      missing.push(key);
      return;
    }

    if (srcRow !== destRow) {
      sheet.moveRows(sheet.getRange(srcRow, 1, 1, maxCols), destRow);
      updateIndexMapAfterMove_(rowByTerm, srcRow, destRow);
    }
    moved.push(key);
    destRow += 1;
  });

  return [
    `Reordered "${labelForMessages}" by term_name.`,
    moved.length ? `Moved to top (in order): ${moved.join(", ")}` : "No rows moved.",
    missing.length ? `Missing (ignored): ${missing.join(", ")}` : "",
  ].filter(Boolean).join("\n");
}

function findBestHeaderRow_(sheet, desiredHeaders, maxRowsToScan) {
  const lastRow = Math.min(sheet.getLastRow(), maxRowsToScan);
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return 1;

  const desiredSet = {};
  desiredHeaders.forEach(h => {
    const key = (h || "").toString().trim();
    if (key) desiredSet[key] = true;
  });

  let bestRow = 1;
  let bestScore = -1;

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  for (let r = 0; r < values.length; r++) {
    let score = 0;
    for (let c = 0; c < values[r].length; c++) {
      const cell = (values[r][c] || "").toString().trim();
      if (desiredSet[cell]) score += 1;
    }
    if (score > bestScore) {
      bestScore = score;
      bestRow = r + 1;
    }
  }

  return bestRow;
}

/**
 * Updates an index map (key -> 1-based index) after a move operation.
 * Works for both columns and rows.
 */
function updateIndexMapAfterMove_(indexMap, srcIndex, destIndex) {
  if (srcIndex === destIndex) return;

  Object.keys(indexMap).forEach(k => {
    let idx = indexMap[k];
    if (idx === srcIndex) {
      idx = destIndex;
    } else if (destIndex < srcIndex) {
      // Moving left/up: items in [destIndex, srcIndex) shift right/down by 1
      if (idx >= destIndex && idx < srcIndex) idx += 1;
    } else {
      // Moving right/down: items in (srcIndex, destIndex] shift left/up by 1
      if (idx > srcIndex && idx <= destIndex) idx -= 1;
    }
    indexMap[k] = idx;
  });
}
```

4. Click the disk icon or `File > Save` to save the script.
5. Close the Apps Script editor.

### Preparing Your Data for the Ocean DNA Explorer (ODE)
<div align="center">
  <img src="src/helpers/node_logo_light_mode.svg" alt="Ocean DNA Explorer Logo" width="200">
</div>

For submission to the [Ocean DNA Explorer](https://www.oceandnaexplorer.org/) and to [edna2obis](https://github.com/aomlomics/edna2obis), you will need to download your data sheets (once you have filled them with data) as TSV files. The Google Apps Script you'll add to your sheet includes a tool to make this easy:

**Steps to Download Your Data:**
1.  After adding the Apps Script (see instructions below), a new menu will appear in your Google Sheet called **FAIReSheets Tools**.
2.  Click **FAIReSheets Tools > Download all sheets as TSV**.
3.  A confirmation dialog will appear showing the folder name that will be created.
4.  Click "Yes" to proceed. The script will create a new folder in your Google Drive home directory (e.g., `FAIRe_NOAA_YourProject_20241112_TSVs_20241112_1430`) and save all sheets as TSV files there.

This will download all your sheets as submission-ready TSV files. The Apps Script also provides helpful data validation features that run automatically when you edit your sheet, helping you catch common errors before submission.

## Disclaimer
This repository is a scientific product and is not official communication of the National Oceanic and Atmospheric Administration, or the United States Department of Commerce. All NOAA GitHub project code is provided on an 'as is' basis and the user assumes responsibility for its use. Any claims against the Department of Commerce or Department of Commerce bureaus stemming from the use of this GitHub project will be governed by all applicable Federal law. Any reference to specific commercial products, processes, or services by service mark, trademark, manufacturer, or otherwise, does not constitute or imply their endorsement, recommendation or favoring by the Department of Commerce. The Department of Commerce seal and logo, or the seal and logo of a DOC bureau, shall not be used in any manner to imply endorsement of any commercial product or activity by DOC or the United States Government.