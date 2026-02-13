"""
Gmail API Authentication Module
Handles OAuth2 authentication for Gmail API access
"""

import os
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.pickle
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.send'
]

def confirm_auth():
    """
    Authenticate with Gmail API and return the service object.
    
    Returns:
        service: Authorized Gmail API service instance
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens
    script_dir = os.path.dirname(os.path.abspath(__file__))
    token_path = os.path.join(script_dir, 'token.pickle')
    credentials_path = os.path.join(script_dir, 'credentials.json')
    
    # Check if token.pickle exists
    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing Gmail access token...")
            creds.refresh(Request())
        else:
            if not os.path.exists(credentials_path):
                raise FileNotFoundError(
                    f"credentials.json not found. Please download it from Google Cloud Console.\n"
                    f"Expected location: {os.path.abspath(credentials_path)}"
                )
            print("Starting Gmail authentication flow...")
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)
        print("Gmail authentication successful!")
    
    # Build and return the Gmail service
    service = build('gmail', 'v1', credentials=creds)
    return service


if __name__ == "__main__":
    # Test authentication
    print("Testing Gmail authentication...")
    service = confirm_auth()
    print("Authentication successful!")
    
    # Test by getting user profile
    profile = service.users().getProfile(userId='me').execute()
    print(f"Authenticated as: {profile.get('emailAddress')}")
