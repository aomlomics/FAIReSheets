"""
Main entry point for FAIReSheets and FAIRe2NOAA tools.
"""

import os
import sys
import yaml
import warnings

from dotenv import load_dotenv
from rich.console import Console
import pyfiglet

from src.auth import authenticate
from src.FAIReSheets import FAIReSheets
from src.FAIRe2NOAA import FAIRe2NOAA

# Suppress specific openpyxl warnings about data validation in the console
warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    message="Data Validation extension is not supported and will be removed",
)

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Initialize rich console
console = Console()

# Display welcome message with ASCII art banner
banner = pyfiglet.figlet_format("FAIReSheets")
console.print()
console.print(banner, style="bold cyan")
console.print(
    "FAIReSheets generates FAIR eDNA metadata templates in Google Sheets.",
    style="bold",
)
console.print(
    "It can create standard FAIR eDNA metadata templates and "
    "Ocean DNA Explorer-compatible FAIR eDNA metadata templates.\n"
)
console.print(
    "First-time users will be prompted to authenticate with Google.\n"
    "NOTE: You must be on the approved users list to use this tool.\n"
    "To request access, email bayden.willms@noaa.gov\n"
)

def main():
    """Main function to run FAIReSheets."""
    # Load environment variables
    load_dotenv()
    
    # Check if .env file exists, if not create it
    if not os.path.exists('.env'):
        with open('.env', 'w') as f:
            f.write('SPREADSHEET_ID=your_spreadsheet_id_here\n')
            f.write('GIST_URL=your_gist_url_here\n')
        console.print(
            "\nCreated .env file. Please edit it with your spreadsheet ID and Gist URL.",
            style="bold yellow",
        )
        return

    # Authenticate with Google
    try:
        client = authenticate()
    except Exception as e:
        console.print(f"\nAuthentication failed: {e}", style="bold red")
        return

    try:
        # Load the configuration files
        config_path = os.path.join(current_dir, 'config.yaml')
        noaa_config_path = os.path.join(current_dir, 'NOAA_config.yaml')

        with open(config_path, 'r') as file:
            config = yaml.safe_load(file)
        
        with open(noaa_config_path, 'r') as file:
            noaa_config = yaml.safe_load(file)

        # Get parameters from config
        run_noaa_formatting = noaa_config.get('run_noaa_formatting', False)
        project_id = config.get('project_id', 'default_project')
        req_lev = config.get('req_lev', ['M', 'HR', 'R', 'O'])
        sample_type = config.get('sample_type', None)
        assay_type = config.get('assay_type', None)
        assay_name = config.get('assay_name', None)
        projectMetadata_user = config.get('projectMetadata_user', None)
        sampleMetadata_user = config.get('sampleMetadata_user', None)
        experimentRunMetadata_user = config.get('experimentRunMetadata_user', None)
        input_dir = os.path.join(current_dir, 'input')

        # Step 1: Generate FAIRe template (spinner)
        console.print("\nStep 1: Generating FAIRe template...", style="bold")
        with console.status(
            "[bold cyan]Generating FAIRe template...[/bold cyan]", spinner="dots"
        ):
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
                client=client,
            )
        
        if run_noaa_formatting:
            # Step 2: Convert to Ocean DNA Explorer format (spinner)
            console.print(
                "\nStep 2: Converting to Ocean DNA Explorer format...", style="bold"
            )
            with console.status(
                "[bold cyan]Converting to Ocean DNA Explorer format...[/bold cyan]",
                spinner="dots",
            ):
                FAIRe2NOAA(client=client, project_id=project_id)
            
            console.print(
                "\nAll done! Your Ocean DNA Explorer-compatible template is ready!",
                style="bold green",
            )
        else:
            console.print(
                "\nAll done! Your FAIReSheets template is ready!",
                style="bold green",
            )
        
    except Exception as e:
        console.print(f"\nError: {e}", style="bold red")

if __name__ == "__main__":
    main() 