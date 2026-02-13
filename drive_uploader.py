"""
Google Drive Upload Helper
Handles authentication and file uploads to Google Drive folders.
"""

import os
import re
import pickle
import logging
from datetime import datetime
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Set up logging
log_file = 'drive_uploads.log'
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()  # Also log to console
    ]
)
logger = logging.getLogger(__name__)

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
            logger.info("Refreshing expired Drive credentials")
            creds.refresh(Request())
        else:
            logger.info("Initiating new Drive authentication")
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for future use
        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)
        logger.info("Drive credentials saved")
    
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
        logger.info(f"Using existing date subfolder: {date_str} (ID: {existing[0]['id']})")
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
    
    folder_id = folder.get('id')
    logger.info(f"Created new date subfolder: {date_str} (ID: {folder_id})")
    return folder_id


def extract_date_from_filename(filename):
    """
    Extract a date (YYYY-MM-DD) from a filename.
    E.g., 'Report-Custom Report-2026-02-09.xlsx' -> '2026-02-09'
    """
    match = re.search(r'(\d{4}-\d{2}-\d{2})', filename)
    return match.group(1) if match else None


def format_file_size(bytes_size):
    """Format bytes to human-readable size."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if bytes_size < 1024.0:
            return f"{bytes_size:.1f} {unit}"
        bytes_size /= 1024.0
    return f"{bytes_size:.1f} TB"


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
        if files:
            logger.debug(f"Found existing file: {filename} (ID: {files[0]['id']})")
        return files[0]['id'] if files else None
    except Exception as e:
        logger.warning(f"Could not check for existing file {filename}: {e}")
        print(f"Warning: Could not check for existing file: {e}")
        return None


def upload_file_to_drive(file_path, folder_name, status_callback=None, target_folder_id=None):
    """
    Upload a file to the specified Google Drive folder with progress tracking.
    
    Args:
        file_path: Local path to file to upload
        folder_name: Name of the folder (e.g., "Andy & Greg")
        status_callback: Optional callback function for status updates
        target_folder_id: If provided, upload directly to this folder ID
    
    Returns:
        True if successful, False otherwise
    """
    file_name = os.path.basename(file_path)
    
    try:
        # Get folder ID
        folder_id = target_folder_id or DRIVE_FOLDERS.get(folder_name)
        if not folder_id or (isinstance(folder_id, str) and folder_id.startswith("REPLACE_WITH")):
            msg = f"❌ Drive folder not configured for '{folder_name}'. Please set folder ID in DriveUploader.py"
            logger.error(f"Drive folder not configured: {folder_name}")
            print(msg)
            if status_callback:
                status_callback(msg)
            return False
        
        # Get file size for progress tracking
        file_size = os.path.getsize(file_path)
        file_size_str = format_file_size(file_size)
        
        # Get Drive service
        service = get_drive_service()
        
        # Check if file already exists in this folder
        existing_file_id = find_existing_file(service, folder_id, file_name)
        
        # Use chunked upload for files > 5MB, simple upload for smaller files
        chunksize = 5 * 1024 * 1024 if file_size > 5 * 1024 * 1024 else -1
        media = MediaFileUpload(file_path, resumable=True, chunksize=chunksize)
        
        if existing_file_id:
            # Update existing file
            msg = f"🔄 Updating existing file: {file_name} ({file_size_str})..."
            logger.info(f"Updating existing file: {file_name} in folder {folder_name} (ID: {existing_file_id})")
            print(msg)
            if status_callback:
                status_callback(msg)
            
            request = service.files().update(
                fileId=existing_file_id,
                media_body=media,
                fields='id, name, webViewLink',
                supportsAllDrives=True
            )
        else:
            # Create new file
            msg = f"📤 Uploading {file_name} ({file_size_str}) to Drive..."
            logger.info(f"Uploading new file: {file_name} to folder {folder_name} ({file_size_str})")
            print(msg)
            if status_callback:
                status_callback(msg)
            
            file_metadata = {
                'name': file_name,
                'parents': [folder_id]
            }
            
            request = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id, name, webViewLink',
                supportsAllDrives=True
            )
        
        # Execute upload with progress tracking for large files
        response = None
        last_progress = 0
        
        while response is None:
            status, response = request.next_chunk()
            if status:
                progress = int(status.progress() * 100)
                # Only update if progress changed by 10% or more (reduce noise)
                if progress >= last_progress + 10:
                    bytes_uploaded = int(status.progress() * file_size)
                    progress_msg = f"   Progress: {progress}% ({format_file_size(bytes_uploaded)} / {file_size_str})"
                    print(progress_msg)
                    if status_callback:
                        status_callback(progress_msg)
                    last_progress = progress
        
        uploaded_file = response
        action = "Updated" if existing_file_id else "Uploaded"
        msg = f"✅ {action}: {file_name} (ID: {uploaded_file.get('id')})"
        logger.info(f"{action} successfully: {file_name} (ID: {uploaded_file.get('id')})")
        print(msg)
        if status_callback:
            status_callback(msg)
        
        return True
        
    except Exception as e:
        msg = f"❌ Error uploading {file_name} to Drive: {e}"
        logger.error(f"Upload failed for {file_name}: {e}", exc_info=True)
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
    logger.info(f"Starting upload batch to '{folder_name}' from {folder_path}")
    
    if not os.path.exists(folder_path):
        msg = f"❌ Folder not found: {folder_path}"
        logger.error(f"Folder not found: {folder_path}")
        print(msg)
        if status_callback:
            status_callback(msg)
        return 0
    
    # Get all .xlsx files
    xlsx_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    if not xlsx_files:
        msg = f"ℹ️ No Excel files found in {folder_path}"
        logger.info(f"No Excel files found in {folder_path}")
        print(msg)
        if status_callback:
            status_callback(msg)
        return 0
    
    msg = f"📂 Found {len(xlsx_files)} file(s) to upload to Drive"
    logger.info(f"Found {len(xlsx_files)} Excel files to upload")
    print(msg)
    if status_callback:
        status_callback(msg)
    
    # Get the parent folder ID
    parent_folder_id = DRIVE_FOLDERS.get(folder_name)
    if not parent_folder_id:
        msg = f"❌ Drive folder not configured for '{folder_name}'"
        logger.error(f"Drive folder not configured: {folder_name}")
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
            logger.warning(f"Could not create date subfolder: {e}")
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
            error_msg = f"Error processing {xlsx_file}: {e}"
            logger.error(error_msg, exc_info=True)
            print(error_msg)
            if status_callback:
                status_callback(error_msg)
    
    # Summary message
    summary_parts = []
    if new_count > 0:
        summary_parts.append(f"{new_count} new")
    if update_count > 0:
        summary_parts.append(f"{update_count} updated")
    
    summary = " and ".join(summary_parts) if summary_parts else "0"
    msg = f"✅ Successfully uploaded {success_count}/{len(xlsx_files)} files ({summary})"
    logger.info(f"Upload batch complete: {success_count}/{len(xlsx_files)} files - {summary}")
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
