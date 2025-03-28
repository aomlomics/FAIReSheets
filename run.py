import os
import sys
import yaml
import warnings

# Suppress specific openpyxl warnings about data validation in the console
warnings.filterwarnings("ignore", category=UserWarning, 
                       message="Data Validation extension is not supported and will be removed")

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Import authentication from our new auth module
from src.auth import authenticate

# Display welcome message
print("\n===================================================")
print("Welcome to FAIReSheets - FAIR eDNA Template Generator")
print("===================================================")
print("This tool generates FAIR eDNA data templates in Google Sheets.")
print("First-time users will be prompted to authenticate with Google.")
print("NOTE: You must be on the approved users list to use this tool.")
print("To request access, email bayden.willms@noaa.gov\n")

# First, check if .env file exists, and create a template if it doesn't
env_file = os.path.join(current_dir, '.env')
if not os.path.exists(env_file):
    print("Creating template .env file...")
    with open(env_file, 'w') as f:
        f.write("# Add your Google Sheet ID below\n")
        f.write("SPREADSHEET_ID=\n\n")
        f.write("# The URL below will be provided when you email bayden.willms@noaa.gov\n")
        f.write("GIST_URL=\n")
    print(f"Created template .env file at {env_file}")
    print("Please edit this file to add your Google Sheet ID and Gist URL (if provided).\n")

# Authenticate with Google
client = authenticate()

# Try to get spreadsheet ID from .env if it exists
from dotenv import load_dotenv
load_dotenv()
spreadsheet_id = os.getenv("SPREADSHEET_ID")

# If not found in .env, ask the user
if not spreadsheet_id:
    print("\n===================================================")
    print("Google Sheet Setup")
    print("===================================================")
    print("You need to provide a Google Sheet ID where the template will be created.")
    print("You can find this in the URL of your Google Sheet:")
    print("https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")
    
    spreadsheet_id = input("\nPlease enter your Google Sheet ID: ").strip()
    
    # Save to .env for future use
    with open(env_file, 'r') as f:
        env_contents = f.readlines()
    
    with open(env_file, 'w') as f:
        for line in env_contents:
            if line.startswith("SPREADSHEET_ID="):
                f.write(f"SPREADSHEET_ID={spreadsheet_id}\n")
            else:
                f.write(line)
    
    print(f"Saved Google Sheet ID to {env_file}")

# Set environment variables for the application
os.environ["SPREADSHEET_ID"] = spreadsheet_id

# Import the function directly from the file in the src directory
from src.FAIReSheets import FAIReSheets

# Load the configuration
try:
    config_path = os.path.join(current_dir, 'config.yaml')
    with open(config_path, 'r') as file:
        config = yaml.safe_load(file)
except Exception as e:
    print(f"Error loading config: {e}")
    sys.exit(1)

# Default parameters for the function
project_id = config.get('project_id', 'default_project')
req_lev = config.get('req_lev', ['M', 'HR', 'R', 'O'])
sample_type = config.get('sample_type', None)
assay_type = config.get('assay_type', None)
assay_name = config.get('assay_name', None)
projectMetadata_user = config.get('projectMetadata_user', None)
sampleMetadata_user = config.get('sampleMetadata_user', None)
experimentRunMetadata_user = config.get('experimentRunMetadata_user', None)
input_dir = os.path.join(current_dir, 'input')

# Run the FAIReSheets function with the configured parameters
print("\n===================================================")
print(f"Generating FAIR eDNA template for project: {project_id}")
print("===================================================")
print("Starting template generation process. This may take up to 2 minutes.")
print("Please be patient and do not close the application during this process. If you think the code is frozen, open the Google Sheet in your browser and watch the progress.\n")

try:
    FAIReSheets(
        req_lev=req_lev,
        sample_type=sample_type,
        assay_type=assay_type,
        project_id=project_id,
        assay_name=assay_name,
        projectMetadata_user=projectMetadata_user,
        sampleMetadata_user=sampleMetadata_user,
        experimentRunMetadata_user=experimentRunMetadata_user,
        input_dir=input_dir,
        client=client  # Pass the authenticated client
    )
    print("\n===================================================")
    print("FAIReSheets completed successfully!")
    print("===================================================")
    print(f"Your Google Sheet has been updated: https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit")
except Exception as e:
    print(f"\nERROR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1) 