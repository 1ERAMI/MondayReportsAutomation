"""
Google Drive Upload Helper
Handles authentication and file uploads to Google Drive folders.
"""

import os
import re
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Google Drive API scope
SCOPES = ['https://www.googleapis.com/auth/drive']

# Folder mappings - replace with your actual Google Drive folder IDs
DRIVE_FOLDERS = {
    "Andy & Greg": "1qOnyoZl_lbWkUGk8r6iMroy8ZTKom91E",
    "Cameron Flatirons": "11V0Ity9HLncxsOkS-yedctUWRO6oiLFF",
    "Cameron & Crump": "1GHYmpl983zq264sZnYj-iR-M_1-w-5LJ",
    "Malissa": "1FroHjovKsopPtRTlr_LPwai7y4isX-DY"
}


def get_drive_service():
    """
    Authenticate and return Google Drive service.
    Uses token_drive.pickle for cached credentials.
    """
    creds = None
    token_file = 'token_drive.pickle'
    
    # Check if token exists
    if os.path.exists(token_file):
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)
    
    # If no valid credentials, authenticate
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for future use
        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)
    
    return build('drive', 'v3', credentials=creds)


def get_or_create_date_subfolder(service, parent_folder_id, date_str):
    """
    Find or create a date subfolder inside a parent Drive folder.
    
    Args:
        service: Google Drive API service
        parent_folder_id: ID of the parent folder
        date_str: Date string to use as subfolder name (e.g., "2026-02-09")
    
    Returns:
        Folder ID of the date subfolder
    """
    # Check if subfolder already exists
    query = (
        f"'{parent_folder_id}' in parents "
        f"and name = '{date_str}' "
        f"and mimeType = 'application/vnd.google-apps.folder' "
        f"and trashed = false"
    )
    results = service.files().list(
        q=query,
        fields='files(id, name)',
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    
    existing = results.get('files', [])
    if existing:
        return existing[0]['id']
    
    # Create the subfolder
    metadata = {
        'name': date_str,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_folder_id]
    }
    folder = service.files().create(
        body=metadata,
        fields='id, name',
        supportsAllDrives=True
    ).execute()
    
    return folder.get('id')


def extract_date_from_filename(filename):
    """
    Extract a date (YYYY-MM-DD) from a filename.
    E.g., 'Report-Custom Report-2026-02-09.xlsx' -> '2026-02-09'
    """
    match = re.search(r'(\d{4}-\d{2}-\d{2})', filename)
    return match.group(1) if match else None


def find_existing_file(service, folder_id, filename):
    """
    Check if a file with the same name already exists in the folder.
    
    Args:
        service: Google Drive API service
        folder_id: ID of the folder to search in
        filename: Name of the file to search for
    
    Returns:
        File ID if found, None otherwise
    """
    # Escape single quotes in filename for query
    escaped_filename = filename.replace("'", "\\'")
    
    query = (
        f"'{folder_id}' in parents "
        f"and name = '{escaped_filename}' "
        f"and trashed = false"
    )
    
    try:
        results = service.files().list(
            q=query,
            fields='files(id, name, modifiedTime)',
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        
        files = results.get('files', [])
        return files[0]['id'] if files else None
    except Exception as e:
        print(f"Warning: Could not check for existing file: {e}")
        return None


def upload_file_to_drive(file_path, folder_name, status_callback=None, target_folder_id=None):
    """
    Upload a file to the specified Google Drive folder.
    
    Args:
        file_path: Local path to file to upload
        folder_name: Name of the folder (e.g., "Andy & Greg")
        status_callback: Optional callback function for status updates
        target_folder_id: If provided, upload directly to this folder ID
    
    Returns:
        True if successful, False otherwise
    """
    try:
        # Get folder ID
        folder_id = target_folder_id or DRIVE_FOLDERS.get(folder_name)
        if not folder_id or (isinstance(folder_id, str) and folder_id.startswith("REPLACE_WITH")):
            msg = f"❌ Drive folder not configured for '{folder_name}'. Please set folder ID in DriveUploader.py"
            print(msg)
            if status_callback:
                status_callback(msg)
            return False
        
        # Get Drive service
        service = get_drive_service()
        
        # File metadata
        file_name = os.path.basename(file_path)
        
        # Check if file already exists in this folder
        existing_file_id = find_existing_file(service, folder_id, file_name)
        
        media = MediaFileUpload(file_path, resumable=True)
        
        if existing_file_id:
            # Update existing file
            msg = f"🔄 Updating existing file: {file_name}..."
            print(msg)
            if status_callback:
                status_callback(msg)
            
            uploaded_file = service.files().update(
                fileId=existing_file_id,
                media_body=media,
                fields='id, name, webViewLink',
                supportsAllDrives=True
            ).execute()
            
            msg = f"✅ Updated: {file_name} (ID: {uploaded_file.get('id')})"
            print(msg)
            if status_callback:
                status_callback(msg)
        else:
            # Create new file
            msg = f"📤 Uploading {file_name} to Drive..."
            print(msg)
            if status_callback:
                status_callback(msg)
            
            file_metadata = {
                'name': file_name,
                'parents': [folder_id]
            }
            
            uploaded_file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id, name, webViewLink',
                supportsAllDrives=True
            ).execute()
            
            msg = f"✅ Uploaded: {file_name} (ID: {uploaded_file.get('id')})"
            print(msg)
            if status_callback:
                status_callback(msg)
        
        return True
        
    except Exception as e:
        file_name = os.path.basename(file_path)
        msg = f"❌ Error uploading {file_name} to Drive: {e}"
        print(msg)
        if status_callback:
            status_callback(msg)
        return False


def upload_folder_to_drive(folder_path, folder_name, status_callback=None):
    """
    Upload all .xlsx files from a folder to Google Drive.
    
    Args:
        folder_path: Local folder containing files to upload
        folder_name: Name of the Drive folder (e.g., "Andy & Greg")
        status_callback: Optional callback function for status updates
    
    Returns:
        Number of files successfully uploaded
    """
    if not os.path.exists(folder_path):
        msg = f"❌ Folder not found: {folder_path}"
        print(msg)
        if status_callback:
            status_callback(msg)
        return 0
    
    # Get all .xlsx files
    xlsx_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    if not xlsx_files:
        msg = f"ℹ️ No Excel files found in {folder_path}"
        print(msg)
        if status_callback:
            status_callback(msg)
        return 0
    
    msg = f"📂 Found {len(xlsx_files)} file(s) to upload to Drive"
    print(msg)
    if status_callback:
        status_callback(msg)
    
    # Get the parent folder ID
    parent_folder_id = DRIVE_FOLDERS.get(folder_name)
    if not parent_folder_id:
        msg = f"❌ Drive folder not configured for '{folder_name}'"
        print(msg)
        if status_callback:
            status_callback(msg)
        return 0
    
    # Extract date from first file to create date subfolder
    report_date = None
    for xlsx_file in xlsx_files:
        report_date = extract_date_from_filename(xlsx_file)
        if report_date:
            break
    
    # Create date subfolder if we found a date
    target_folder_id = parent_folder_id
    if report_date:
        try:
            service = get_drive_service()
            target_folder_id = get_or_create_date_subfolder(service, parent_folder_id, report_date)
            msg = f"📁 Uploading to subfolder: {report_date}"
            print(msg)
            if status_callback:
                status_callback(msg)
        except Exception as e:
            msg = f"⚠️ Could not create date subfolder, uploading to main folder: {e}"
            print(msg)
            if status_callback:
                status_callback(msg)
    
    # Upload each file
    success_count = 0
    update_count = 0
    new_count = 0
    
    for xlsx_file in xlsx_files:
        file_path = os.path.join(folder_path, xlsx_file)
        
        # Check if file exists before uploading
        try:
            service = get_drive_service()
            existing = find_existing_file(service, target_folder_id, xlsx_file)
            
            if upload_file_to_drive(file_path, folder_name, status_callback, target_folder_id=target_folder_id):
                success_count += 1
                if existing:
                    update_count += 1
                else:
                    new_count += 1
        except Exception as e:
            print(f"Error processing {xlsx_file}: {e}")
            if status_callback:
                status_callback(f"Error processing {xlsx_file}: {e}")
    
    # Summary message
    summary_parts = []
    if new_count > 0:
        summary_parts.append(f"{new_count} new")
    if update_count > 0:
        summary_parts.append(f"{update_count} updated")
    
    summary = " and ".join(summary_parts) if summary_parts else "0"
    msg = f"✅ Successfully uploaded {success_count}/{len(xlsx_files)} files ({summary})"
    print(msg)
    if status_callback:
        status_callback(msg)
    
    return success_count


def setup_drive_folders():
    """
    Helper function to display instructions for setting up Drive folders.
    Run this once to get setup instructions.
    """
    print("=" * 60)
    print("GOOGLE DRIVE FOLDER SETUP")
    print("=" * 60)
    print("\n1. Create folders in Google Drive for each report type:")
    print("   - Andy & Greg Reports")
    print("   - Cameron Flatirons Reports")
    print("   - Cameron & Crump Reports")
    print("   - Malissa Reports")
    print("\n2. Open each folder in your browser")
    print("3. Copy the folder ID from the URL:")
    print("   https://drive.google.com/drive/folders/FOLDER_ID_HERE")
    print("\n4. Update DRIVE_FOLDERS dictionary in DriveUploader.py")
    print("=" * 60)


if __name__ == "__main__":
    # Run this to see setup instructions
    setup_drive_folders()
