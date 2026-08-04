"""
Authentication module for FAIReSheets.

This module handles OAuth authentication for Google Sheets API access.
"""

import os
import json
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Define scopes needed for Google Sheets
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Desktop OAuth clients cannot keep credentials confidential. Google treats this
# client secret as part of the installed-app configuration, not as a secure secret.
OAUTH_CLIENT_CONFIG = {
    "installed": {
        "client_id": (
            "664841099827-g82hj56glafrmtq0sco3b6rl72i67tr1"
            ".apps.googleusercontent.com"
        ),
        "project_id": "fairesheets",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_secret": "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",
        "redirect_uris": ["http://localhost"],
    }
}

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
    
    # Check for token
    credentials = None
    if os.path.exists(token_file):
        try:
            with open(token_file, encoding="utf-8") as file:
                credentials = Credentials.from_authorized_user_info(
                    json.load(file), SCOPES
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
            # Start the installed-app OAuth flow with PKCE
            flow = InstalledAppFlow.from_client_config(
                OAUTH_CLIENT_CONFIG,
                SCOPES,
                autogenerate_code_verifier=True,
            )
            
            print("\n============================================")
            print("Starting Google Authentication")
            print("============================================")
            print("A browser window will open for you to sign in with your Google account.")
            print("\nWaiting for browser authentication...")
            
            credentials = flow.run_local_server(
                host="127.0.0.1",
                port=0,
                authorization_prompt_message=(
                    "Your browser will open to authorize Google Sheets access."
                ),
                success_message=(
                    "Authentication complete. You may close this window and return to the app."
                ),
                open_browser=True,
            )
            
            # Save token
            with open(token_file, "w", encoding="utf-8") as f:
                f.write(credentials.to_json())
            print("Authentication successful! Token saved for future use.")
    
    client = gspread.authorize(credentials)
    print("Successfully authenticated with Google!")
    return client 