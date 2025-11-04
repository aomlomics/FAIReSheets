<div align="center">
  <img src="src/helpers/banner_fairesheets.png" alt="FAIReSheets Banner" width="800">
</div>

FAIReSheets converts the FAIR eDNA ([FAIRe](https://fair-edna.github.io/index.html)) data checklist to customizable Google Sheets templates. FAIReSheets can be run in one of 2 modes:
1. **FAIRe eDNA:** Generate FAIRe eDNA data templates from the FAIRe checklist
2. **ODE-ready (NOAA):** Generate ODE-ready (NOAA) data templates, ready for submission to the [Ocean DNA Explorer](https://www.oceandnaexplorer.org/), and can be used as input to [edna2obis](https://github.com/aomlomics/edna2obis), a data pipeline for submission to [GBIF](https://www.gbif.org/) and [OBIS](https://obis.org/).

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

## Optional (recommended)

To automatically track modifications in your Google Sheet, you can use the following Apps Script. This script updates the modification timestamp and user email in the README sheet whenever an edit is made to any sheet except "README" and "Drop-down values".

### Steps to Add the Apps Script

1. Open your Google Sheet.
2. Click on `Extensions` in the menu, then select `Apps Script`.
3. Delete any code in the script editor and copy-paste the following code:

```javascript
function onEdit(e) {
  // Get the edited range and sheet
  var range = e.range;
  var sheet = range.getSheet();
  var spreadsheet = sheet.getParent();
  var sheetName = sheet.getName();
  
  // First handle the timestamp update (original functionality)
  // Skip README and Drop-down values sheets for timestamp update
  if (sheetName !== "README" && sheetName !== "Drop-down values") {
    var readmeSheet = spreadsheet.getSheetByName("README");
    if (readmeSheet) {
      // Find the Sheets in this Google Sheet section
      var data = readmeSheet.getDataRange().getValues();
      var timestampRowStart = -1;
      
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === "Sheets in this Google Sheet:") {
          timestampRowStart = i + 2; // +2 to skip the header row
          break;
        }
      }
      
      if (timestampRowStart !== -1) {
        // Find the row for the current sheet
        var sheetRow = -1;
        for (var i = timestampRowStart; i < data.length; i++) {
          if (data[i][0] === sheetName) {
            sheetRow = i + 1; // +1 because arrays are 0-indexed but sheets are 1-indexed
            break;
          }
        }
        
        if (sheetRow !== -1) {
          // Get current time in ISO format
          var now = new Date();
          var timestamp = now.toISOString();
          
          // Update the timestamp and email
          readmeSheet.getRange(sheetRow, 2).setValue(timestamp);
          readmeSheet.getRange(sheetRow, 3).setValue(Session.getActiveUser().getEmail());
        }
      }
    }
  }
  
  // Then handle the validation checks (new functionality)
  // Skip validation for README and Drop-down values sheets
  if (sheetName === "README" || sheetName === "Drop-down values") {
    return;
  }
  
  // Get all sheets
  var projectMetadataSheet = spreadsheet.getSheetByName("projectMetadata");
  var analysisMetadataSheets = spreadsheet.getSheets().filter(s => 
    s.getName().startsWith("analysisMetadata_"));
  
  if (!projectMetadataSheet || analysisMetadataSheets.length === 0) {
    return; // Exit if required sheets don't exist
  }
  
  // Get project_id from projectMetadata
  var projectIdCell = findCellByValue(projectMetadataSheet, "project_id");
  if (!projectIdCell) return;
  
  var projectId = projectMetadataSheet.getRange(projectIdCell.row, projectIdCell.col + 1).getValue();
  
  // Get assay_name values from projectMetadata
  var assayNameCell = findCellByValue(projectMetadataSheet, "assay_name");
  if (!assayNameCell) return;
  
  var assayNames = projectMetadataSheet.getRange(assayNameCell.row, assayNameCell.col + 1)
    .getValue()
    .split("|")
    .map(name => name.trim());
  
  // Clear only error-related formatting
  clearErrorFormatting(spreadsheet);
  
  // Verify project_id in all analysisMetadata sheets
  var hasProjectIdError = false;
  analysisMetadataSheets.forEach(analysisSheet => {
    var analysisProjectIdCell = findCellByValue(analysisSheet, "project_id");
    if (analysisProjectIdCell) {
      var analysisProjectId = analysisSheet.getRange(analysisProjectIdCell.row, analysisProjectIdCell.col + 1).getValue();
      if (analysisProjectId !== projectId) {
        addErrorFormatting(analysisSheet, analysisProjectIdCell.row, analysisProjectIdCell.col + 1, 
          "Project ID must match the one in projectMetadata sheet");
        hasProjectIdError = true;
      }
    }
  });
  
  // Verify assay_name values
  var hasAssayNameError = false;
  var foundAssayNames = new Set();
  
  analysisMetadataSheets.forEach(analysisSheet => {
    var analysisAssayNameCell = findCellByValue(analysisSheet, "assay_name");
    if (analysisAssayNameCell) {
      var analysisAssayName = analysisSheet.getRange(analysisAssayNameCell.row, analysisAssayNameCell.col + 1).getValue();
      if (!assayNames.includes(analysisAssayName)) {
        addErrorFormatting(analysisSheet, analysisAssayNameCell.row, analysisAssayNameCell.col + 1,
          "Assay name must match one of the values in projectMetadata sheet");
        hasAssayNameError = true;
      } else {
        foundAssayNames.add(analysisAssayName);
      }
    }
  });
  
  // Check if all assay names from projectMetadata are used
  assayNames.forEach(assayName => {
    if (!foundAssayNames.has(assayName)) {
      addErrorFormatting(projectMetadataSheet, assayNameCell.row, assayNameCell.col + 1,
        "Each assay name must be used in at least one analysisMetadata sheet");
      hasAssayNameError = true;
    }
  });

  // Verify analysis_run_name uniqueness within each sheet
  analysisMetadataSheets.forEach(analysisSheet => {
    var analysisRunNameCell = findCellByValue(analysisSheet, "analysis_run_name");
    if (analysisRunNameCell) {
      // Get all values in the column
      var data = analysisSheet.getDataRange().getValues();
      var analysisRunNameCol = analysisRunNameCell.col - 1; // Convert to 0-based index
      var analysisRunNames = new Set();
      
      // Check each row in the column
      for (var i = 0; i < data.length; i++) {
        var value = data[i][analysisRunNameCol];
        if (value && value.trim() !== "") { // Only check non-empty values
          if (analysisRunNames.has(value)) {
            // Found a duplicate
            addErrorFormatting(analysisSheet, i + 1, analysisRunNameCell.col,
              "Duplicate analysis_run_name found in this sheet");
          } else {
            analysisRunNames.add(value);
          }
        }
      }
    }
  });
  var font = "Arial"; // Set your desired font
  var fontSize = 10;  // Set your desired font size
  range.setFontFamily(font);
  range.setFontSize(fontSize);
}

// Helper function to find a cell containing a specific value
function findCellByValue(sheet, searchValue) {
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] === searchValue) {
        return {row: i + 1, col: j + 1};
      }
    }
  }
  return null;
}

// Helper function to add error formatting
function addErrorFormatting(sheet, row, col, message) {
  var cell = sheet.getRange(row, col);
  // Store the current note if it exists
  var currentNote = cell.getNote();
  // Only add the error message if it's not already there
  if (!currentNote.includes(message)) {
    cell.setNote(currentNote + (currentNote ? "\n" : "") + "ERROR: " + message);
  }
  // Add red background only if it's not already an error
  if (cell.getBackground() !== "#ff0000") {
    cell.setBackground("#ff0000"); // Light red background
  }
}

// Helper function to clear only error-related formatting
function clearErrorFormatting(spreadsheet) {
  var sheets = spreadsheet.getSheets();
  sheets.forEach(sheet => {
    var dataRange = sheet.getDataRange();
    var backgrounds = dataRange.getBackgrounds();
    var notes = dataRange.getNotes();
    
    // Clear only error-related formatting
    for (var i = 0; i < backgrounds.length; i++) {
      for (var j = 0; j < backgrounds[i].length; j++) {
        if (backgrounds[i][j] === "#ff0000") { // If it's our error background color
          var cell = sheet.getRange(i + 1, j + 1);
          // Remove only the error message from the note, keep other notes
          var note = notes[i][j];
          if (note) {
            var errorLines = note.split("\n").filter(line => !line.startsWith("ERROR:"));
            cell.setNote(errorLines.join("\n"));
          }
          cell.setBackground(null); // Remove the red background
        }
      }
    }
  });
}
```

4. Click the disk icon or `File > Save` to save the script.
5. Close the Apps Script editor.

Now, whenever you make an edit to any sheet (except "README" and "Drop-down values"), the modification timestamp and your email will be automatically updated in the README sheet.

## Disclaimer
This repository is a scientific product and is not official communication of the National Oceanic and Atmospheric Administration, or the United States Department of Commerce. All NOAA GitHub project code is provided on an 'as is' basis and the user assumes responsibility for its use. Any claims against the Department of Commerce or Department of Commerce bureaus stemming from the use of this GitHub project will be governed by all applicable Federal law. Any reference to specific commercial products, processes, or services by service mark, trademark, manufacturer, or otherwise, does not constitute or imply their endorsement, recommendation or favoring by the Department of Commerce. The Department of Commerce seal and logo, or the seal and logo of a DOC bureau, shall not be used in any manner to imply endorsement of any commercial product or activity by DOC or the United States Government.