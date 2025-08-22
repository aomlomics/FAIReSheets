"""
Main entry point for FAIReSheets and FAIRe2NODE tools.
"""

import os
import sys
import yaml
import warnings
from dotenv import load_dotenv
from src.auth import authenticate
from src.FAIReSheets import FAIReSheets
from src.FAIRe2NODE import FAIRe2NODE

# Suppress specific openpyxl warnings about data validation in the console
warnings.filterwarnings("ignore", category=UserWarning, 
                       message="Data Validation extension is not supported and will be removed")

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Display welcome message
print("\n===================================================")
print("Welcome to FAIReSheets - FAIR eDNA Template Generator")
print("===================================================")
print("This tool generates FAIR eDNA data templates in Google Sheets")
print("and converts them to ODE format.")
print("First-time users will be prompted to authenticate with Google.")
print("NOTE: You must be on the approved users list to use this tool.")
print("To request access, email bayden.willms@noaa.gov\n")

def main():
    """Main function to run FAIReSheets followed by FAIRe2ODE."""
    # Load environment variables
    load_dotenv()
    
    # Check if .env file exists, if not create it
    if not os.path.exists('.env'):
        with open('.env', 'w') as f:
            f.write('SPREADSHEET_ID=your_spreadsheet_id_here\n')
            f.write('GIST_URL=your_gist_url_here\n')
        print("\n⚠️  Created .env file. Please edit it with your spreadsheet ID and Gist URL.")
        return

    # Authenticate with Google
    try:
        client = authenticate()
    except Exception as e:
        print(f"\n❌ Authentication failed: {e}")
        return

    try:
        # Load the configuration
        config_path = os.path.join(current_dir, 'config.yaml')
        with open(config_path, 'r') as file:
            config = yaml.safe_load(file)
            
        # Get parameters from config
        project_id = config.get('project_id', 'default_project')
        req_lev = config.get('req_lev', ['M', 'HR', 'R', 'O'])
        sample_type = config.get('sample_type', None)
        assay_type = config.get('assay_type', None)
        assay_name = config.get('assay_name', None)
        projectMetadata_user = config.get('projectMetadata_user', None)
        sampleMetadata_user = config.get('sampleMetadata_user', None)
        experimentRunMetadata_user = config.get('experimentRunMetadata_user', None)
        input_dir = os.path.join(current_dir, 'input')

        # Step 1: Generate FAIRe template
        print("\n📝 Step 1: Generating FAIRe template...")
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
            client=client
        )
        
        # Step 2: Convert to ODE format
        print("\n🔄 Step 2: Converting to ODE format...")
        FAIRe2NODE(client=client, project_id=project_id)
        
        print("\n✨ All done! Your Ocean DNA Explorer-compatible template is ready!")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")

if __name__ == "__main__":
    main() 