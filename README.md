# FAIReSheets

This project is actively under development. Reach out to bayden.willms@noaa.gov for questions.

FAIReSheets generates the FAIRe eDNA Data Template in Google Sheets, instead of outputting to Microsoft Excel. It replicates the template creation from the FAIRe-ator Repository from Miwa Takahashi and Stephen Formel: https://github.com/FAIR-eDNA/FAIRe-ator/tree/main 

TLDR: Create a blank Google Sheet, configure the Sheets API via the Google Cloud Console, create a service account and JSON credential, and share the Google Sheet with that service account. Add the spreadsheet ID into the .env file of this repository, specify your template parameters in `config.yaml` and run FAIReSheets, either through your IDE or in the terminal via:
```bash
python run.py 
```

## Setup

### Administrator Setup
A Google Drive admin account in your organization must grant permission for the Google Sheets API in Google Cloud Console. Alternatively, you can use a personal Google Drive account. 

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Log in with the account that has administrator permissions.
3. Name the project whatever you'd like.

#### Creating a Service Account
1. Go to API & Services > Credentials
2. Click on "Create Credentials" and select "Service Account".
3. Fill in the service account details (name, ID, description) and click "Create".

#### Assign Roles
1. In the next step, assign the "Editor" role to the service account to grant access to modify Google Sheets.
2. Click "Continue".

#### Create Key
1. Under the service account you created, click "Add Key" and select "JSON".
2. This will download a JSON file with the credentials to your computer. **DO NOT SHARE THIS PUBLICLY**.
3. Copy or Move this JSON file into the FAIReSheets project directory. 

#### Grant Access to Service Account
1. The service account has an email address (e.g., `service-account-name@data-templates.iam.gserviceaccount.com`).
2. **Share the Google Sheets with this email address to grant access.**


### User Setup
This project works on both macOS and Windows. Ensure you have Git and Anaconda installed on your computer:

- [Download Git](https://git-scm.com/downloads)
- [Install Anaconda](https://docs.anaconda.com/anaconda/install/)


#### Clone the GitHub project to your local machine:
```bash
git clone https://github.com/baydenwillms/FAIReSheets.git
cd FAIReSheets
```

#### Configure Anaconda environment using provided `environment.yml` 
```bash
conda env create -f environment.yml 
conda activate fairesheets-env
```
#### Alternatively, you can use your regular Python installation/environment, and pip install the dependencies. There are only a few, for example:
```bash
pip install pandas
```
#### Credentials
As previously mentioned, just Copy or Move your JSON credentials file into the root directory of this repository. 

Then, include the Spreadsheet ID in your .env file, or paste directly into the `run.py` script.



## Run

Before running FAIReSheets, configure the `config.yaml` file given your desired parameters. These parameters determine the structure of your generated FAIRe template.

Finally, run FAIReSheets using: 
```bash
python main.py
```
Or alternatively you can run the `main.py` script in your IDE using the Play button.