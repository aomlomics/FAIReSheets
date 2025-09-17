"""
Authentication module for FAIReSheets.

This module handles OAuth authentication for Google Sheets API access.
"""

import os
import sys
import json
import requests
import webbrowser
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from dotenv import load_dotenv

# Define scopes needed for Google Sheets
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def download_client_secrets():
    """Download client secrets file from GitHub Gist URL specified in .env."""
    current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    client_secrets_path = os.path.join(current_dir, 'client_secrets.json')
    
    # Skip if file already exists
    if os.path.exists(client_secrets_path):
        print("OAuth client secrets file found.")
        return True
    
    # Load environment variables to get the gist URL
    load_dotenv()
    gist_url = os.getenv("GIST_URL")
    
    if not gist_url:
        print("\n============================================")
        print("ERROR: Missing GitHub Gist URL")
        print("============================================")
        print("To use this application, please email bayden.willms@noaa.gov and:")
        print("1. Include your Google account email to be added to the approved users list")
        print("2. Request the Gist URL needed for authentication")
        print("3. Add the provided URL to your .env file as: GIST_URL=https://gist.url.here")
        return False
    
    print("\n============================================")
    print("Downloading OAuth credentials...")
    print("============================================")
    
    try:
        response = requests.get(gist_url)
        
        if response.status_code == 200:
            with open(client_secrets_path, 'w') as f:
                f.write(response.text)
            print("OAuth credentials downloaded successfully.")
            return True
        else:
            print(f"Failed to download credentials. Status code: {response.status_code}")
            print("\nIMPORTANT: To use this application, please email bayden.willms@noaa.gov")
            print("Include your Google account email to be added to the approved users list.")
            return False
    except Exception as e:
        print(f"Error downloading credentials: {e}")
        print("\nIMPORTANT: To use this application, please email bayden.willms@noaa.gov")
        print("Include your Google account email to be added to the approved users list.")
        return False

def authenticate():
    """
    Authenticate with Google using OAuth.
    
    This function checks for an existing token file and uses it if valid.
    If no valid token exists, it initiates the OAuth flow which will
    open a browser window for the user to authenticate with Google.
    
    Returns:
        gspread.Client: Authenticated client for Google Sheets API
    """
    current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    token_file = os.path.join(current_dir, "token.json")
    client_secrets_file = os.path.join(current_dir, "client_secrets.json")
    
    # Check for token
    credentials = None
    if os.path.exists(token_file):
        try:
            credentials = Credentials.from_authorized_user_info(
                json.load(open(token_file)), SCOPES
            )
        except Exception as e:
            print(f"Error loading token: {e}")
            credentials = None
    
    # If no credentials or they're invalid
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            try:
                credentials.refresh(Request())
            except Exception as e:
                print(f"Error refreshing token: {e}")
                credentials = None
        
        if not credentials:
            # Make sure we have client secrets
            if not os.path.exists(client_secrets_file):
                if not download_client_secrets():
                    print("\nCannot proceed without authentication.")
                    print("Please email bayden.willms@noaa.gov to request access.")
                    sys.exit(1)
            
            # Start OAuth flow
            flow = InstalledAppFlow.from_client_secrets_file(
                client_secrets_file, SCOPES
            )
            
            print("\n============================================")
            print("Starting Google Authentication")
            print("============================================")
            print("A browser window will open for you to sign in with your Google account.")
            print("\nIMPORTANT: You must be on the approved users list to authenticate.")
            print("If you get an 'app not verified' message and cannot proceed,")
            print("please email bayden.willms@noaa.gov to request access.")
            print("\nWaiting for browser authentication...")
            
            # Try a few stable localhost ports (helps with corporate proxies and IPv6 issues)
            credentials = None
            last_error = None
            for try_port in [8080, 8888, 8787, 0]:  # 0 = random open port as final fallback
                try:
                    credentials = flow.run_local_server(
                        host="127.0.0.1",
                        port=try_port,
                        authorization_prompt_message=(
                            "Your browser will open to authorize Google Sheets/Drive access."
                        ),
                        success_message=(
                            "Authentication complete. You may close this window and return to the app."
                        ),
                        open_browser=True,
                    )
                    break
                except Exception as e:
                    last_error = e
                    continue

            if credentials is None:
                try:
                    print("\nLocalhost redirect did not succeed. Falling back to console authentication...")
                    print("1) A URL will be printed below.\n2) Open it in a browser, authorize, then paste the code here.")
                    credentials = flow.run_console()
                except Exception as e2:
                    raise RuntimeError(
                        f"Unable to complete authentication via localhost or console. Last localhost error: {last_error}; console error: {e2}"
                    )
            
            # Save token
            with open(token_file, "w") as f:
                f.write(credentials.to_json())
            print("Authentication successful! Token saved for future use.")
    
    client = gspread.authorize(credentials)
    print("Successfully authenticated with Google!")
    return client 