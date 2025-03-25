# FAIReSheets

This project is actively under development. Reach out to bayden.willms@noaa.gov for questions.

FAIReSheets generates the FAIRe eDNA Data Template in Google Sheets, instead of outputting to Microsoft Excel. It replicates the template creation from the FAIRe-ator Repository from Miwa Takahashi and Stephen Formel: https://github.com/FAIR-eDNA/FAIRe-ator/tree/main 

TLDR: Create a blank Google Sheet, configure the `.env` file with your Google Sheet ID and the Git Gist URL, run FAIReSheets, and follow the authentication prompts.

```bash
python run.py 
```

## Setup

### Access Request (Required)
Before using FAIReSheets, you'll need to request access:

1. Email bayden.willms@noaa.gov with the subject "FAIReSheets Access Request"
2. Include your Google account email in the request (the one you'll use to access Google Sheets)
3. You'll receive an email with a Gist URL to add to your `.env` file

### Installation
This project works on both macOS and Windows. Ensure you have Git and either Anaconda or pip installed on your computer:

- [Download Git](https://git-scm.com/downloads)
- [Install Anaconda](https://docs.anaconda.com/anaconda/install/) (recommended)

#### Clone the GitHub project to your local machine:
```bash
git clone https://github.com/baydenwillms/FAIReSheets.git
cd FAIReSheets
```

#### Option 1: Configure Anaconda environment 
```bash
conda env create -f environment.yml 
conda activate fairesheets-env
```

#### Option 2: Install with pip
```bash
pip install -r requirements.txt
```

### Google Sheet Setup
1. Create a new, empty Google Sheet in your account
2. Copy the Spreadsheet ID from the URL:
   - The ID is the long string between `/d/` and `/edit` in the URL
   - For example: `https://docs.google.com/spreadsheets/d/1ABC123XYZ/edit`
   - The ID would be `1ABC123XYZ`

### Configure .env File
Run the script once and it will create a template `.env` file:
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

1. The tool will download the necessary authentication credentials
2. A browser window will open asking you to sign in with your Google account
3. You'll be asked to grant permission to FAIReSheets to access your Google Sheets
4. After granting permission, the tool will save a token for future use
5. Future runs won't require reauthentication

If you see a message saying "Google hasn't verified this app," click "Advanced" and then "Go to FAIReSheets (unsafe)" to proceed. This is normal for specialized tools that haven't gone through Google's verification process.

## Run

Before running FAIReSheets, configure the `config.yaml` file given your desired parameters. These parameters determine the structure of your generated FAIRe template.

Finally, run FAIReSheets using: 
```bash
python run.py
```

Or alternatively you can run the `run.py` script in your IDE using the Play button.