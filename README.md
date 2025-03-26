# FAIReSheets

This project is actively under development. Reach out to bayden.willms@noaa.gov for questions and for credentials to run this code.

FAIReSheets generates the FAIRe eDNA Data Template in Google Sheets. The FAIRe template is a collaborative effort in the eDNA research field to standardize its complicated data and metadata. FAIReSheets replicates the template creation from the [FAIRe-ator Repository](https://github.com/FAIR-eDNA/FAIRe-ator/tree/main) from Dr. Miwa Takahashi and Dr. Stephen Formel, except FAIReSheets outputs the template to Google Sheets rather than a Microsoft Excel spreadsheet. 

TLDR: Email bayden.willms@noaa.gov to be added to the user list and receive the link to the credentials file, create a blank Google Sheet, configure the `.env` file with your Google Sheet ID and the Git Gist URL, run FAIReSheets and follow the authentication workflow.

```bash
python run.py 
```

## Setup

### Access Request (Required)
Before using FAIReSheets, you'll need to request access. This only needs to happen once:

1. Email bayden.willms@noaa.gov
2. Include your Google account email in the request (the one you'll use to access Google Sheets)
3. You'll receive an email with a Gist URL to add to your `.env` file and confirmation that you've been added to the user list.

### Installation
This project works on both macOS and Windows. Ensure you have Git and either Anaconda or pip installed on your computer:

- [Download Git](https://git-scm.com/downloads)
- [Install Anaconda](https://docs.anaconda.com/anaconda/install/) OR [Install pip](https://pip.pypa.io/en/stable/installation/) 

#### Clone the GitHub project to your local machine:
```bash
git clone https://github.com/baydenwillms/FAIReSheets.git
cd FAIReSheets
```

#### Option 1: Install with pip
```bash
pip install -r requirements.txt
```
#### Option 2: Configure Anaconda environment 
```bash
conda env create -f environment.yml 
conda activate fairesheets
```

### Google Sheet Setup
1. Create a new, empty Google Sheet in your account
2. Copy the Spreadsheet ID from the URL:
   - The ID is the long string between `/d/` and `/edit` in the URL
   - For example: `https://docs.google.com/spreadsheets/d/1ABC123XYZ/edit`
   - The ID would be `1ABC123XYZ`

### Configure .env File
`.env` files store sensitive variables necessary for running the script. If these variables are properly stored in the `.env` file, they will not be visible on GitHub if you push your local repository publically. Run the script once and it will create a template `.env` file:
```bash
python run.py
```

Then edit the `.env` file to add:
1. Your Google Sheet ID
2. The Gist URL provided to you by email

The `.env` file should look like:
```
SPREADSHEET_ID=your_spreadsheet_id_here
GIST_URL=https://gist.githubusercontent.com/user/hash/raw/file.json
```

## Authentication

When you run FAIReSheets for the first time, the following will happen:

1. The tool will download the authentication credentials from the Git Gist URL sent to you via email
2. A browser window will open asking you to sign in with your Google account
3. You'll be asked to grant permission to FAIReSheets to access your Google Sheets
4. After granting permission, the tool will save a token for future use
5. Future runs won't require reauthentication

If you see a message saying "Google hasn't verified this app," click "Advanced" and then "Go to FAIReSheets (unsafe)" to proceed. This is normal for specialized tools that haven't gone through Google's verification process.

## Run

Before running FAIReSheets, configure the `config.yaml` file given your desired parameters. These parameters determine the structure of your generated FAIRe template. You can also specify additional terms to add if you have relevant fields in your data that are not included in the FAIRe template.

Run FAIReSheets from the root project directory using: 
```bash
python run.py
```
Or alternatively you can run the `run.py` script in your IDE using the Play button. If you are missing things like a `.env` file, authentication credentials, or your spreadsheet ID to edit, the script will prompt you to add those, and/or create a sample `.env` file for you to edit. 

## Optional (recommended)

To automatically track modifications in your Google Sheet, you can use the following Apps Script. This script updates the modification timestamp and user email in the README sheet whenever an edit is made to any sheet except "README" and "Drop-down values".

### Steps to Add the Apps Script

1. Open your Google Sheet.
2. Click on `Extensions` in the menu, then select `Apps Script`.
3. Delete any code in the script editor and copy-paste the following code:

    ```javascript
    function onEdit(e) {
      var sheet = e.source.getActiveSheet();
      var sheetName = sheet.getName();
      
      // Skip README and Drop-down values sheets
      if (sheetName === "README" || sheetName === "Drop-down values") {
        return;
      }
      
      var readmeSheet = e.source.getSheetByName("README");
      if (!readmeSheet) {
        return;
      }
      
      // Find the Modification Timestamp section
      var data = readmeSheet.getDataRange().getValues();
      var timestampRowStart = -1;
      
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === "Modification Timestamp:") {
          timestampRowStart = i + 2; // +2 to skip the header row
          break;
        }
      }
      
      if (timestampRowStart === -1) {
        return;
      }
      
      // Find the row for the current sheet
      var sheetRow = -1;
      for (var i = timestampRowStart; i < data.length; i++) {
        if (data[i][0] === sheetName) {
          sheetRow = i + 1; // +1 because arrays are 0-indexed but sheets are 1-indexed
          break;
        }
      }
      
      if (sheetRow === -1) {
        return;
      }
      
      // Get current time in ISO format
      var now = new Date();
      var timestamp = now.toISOString(); // This gives format like "2025-01-29T13:49:09.123Z"
      
      // Update the timestamp and email
      readmeSheet.getRange(sheetRow, 2).setValue(timestamp);
      readmeSheet.getRange(sheetRow, 3).setValue(Session.getActiveUser().getEmail());
    }
    ```

4. Click the disk icon or `File > Save` to save the script.
5. Close the Apps Script editor.

Now, whenever you make an edit to any sheet (except "README" and "Drop-down values"), the modification timestamp and your email will be automatically updated in the README sheet.

## Disclaimer
This repository is a scientific product and is not official communication of the National Oceanic and Atmospheric Administration, or the United States Department of Commerce. All NOAA GitHub project code is provided on an 'as is' basis and the user assumes responsibility for its use. Any claims against the Department of Commerce or Department of Commerce bureaus stemming from the use of this GitHub project will be governed by all applicable Federal law. Any reference to specific commercial products, processes, or services by service mark, trademark, manufacturer, or otherwise, does not constitute or imply their endorsement, recommendation or favoring by the Department of Commerce. The Department of Commerce seal and logo, or the seal and logo of a DOC bureau, shall not be used in any manner to imply endorsement of any commercial product or activity by DOC or the United States Government.