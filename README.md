# CIR Monday Reports Automation

Automated system for downloading, processing, and emailing Monday reports from Gmail attachments.

## Overview

This project automates the weekly Monday reports workflow by:
1. **Fetching** unread emails with specific subject lines from Gmail
2. **Downloading** Excel file attachments
3. **Processing** the Excel files (formatting, pivot tables, etc.)
4. **Sending** the processed files via email to designated recipients
5. **Uploading** reports to Google Drive with automatic date-based organization

---

## Quick Start

### Unified UI (Recommended)
Run the unified interface to process any or all report types:
```bash
python MondayReportsUI.py
```
**Features:**
- Select multiple report types (runs sequentially)
- Choose recipients with checkboxes
- Email and/or Google Drive delivery (**Drive is default**)
- Automatic date-based folder organization in Drive
- Real-time upload progress for large files
- Comprehensive activity logging (`drive_uploads.log`)
- Duplicate file handling (smart update vs create)
- Modern dark-themed interface
- Live progress updates

### Individual Scripts
Each report can still be run independently:
```bash
python report_andy_greg.py
python report_cameron_flatirons.py
python report_cameron_crump.py
python report_malissa.py
```

---

## Project Structure

```
IND_Tools/
├── MondayReportsUI.py              # Unified UI for all reports (USE THIS)
├── report_config.py                # Centralized configuration (recipients, filters, paths)
├── report_common.py                # Shared functions (download, process, pivot, email)
├── report_andy_greg.py             # Andy & Greg reports wrapper
├── report_cameron_flatirons.py     # Cameron Flatirons reports wrapper
├── report_cameron_crump.py         # Cameron & Crump reports wrapper
├── report_malissa.py               # Malissa reports wrapper
├── gmail_auth.py                   # Gmail API authentication module
├── drive_uploader.py               # Google Drive upload module
├── xlsx_fixer.py                   # Excel column width fixer module
├── DRIVE_SETUP.md                  # Google Drive setup guide
├── credentials.json                # Google API OAuth credentials
├── token.pickle                    # Cached Gmail authentication token
├── token_drive.pickle              # Cached Drive authentication token
├── requirements.txt                # Python dependencies
└── __pycache__/                    # Python cache files
```

### Architecture

```
MondayReportsUI.py (entry point)
  ├── report_andy_greg.py ──┐
  ├── report_cameron_flatirons.py ──┤
  ├── report_cameron_crump.py ──┤── All delegate to report_common.run_report()
  └── report_malissa.py ──┘
                                ↓
                   report_common.py (shared logic)
                   ├── gmail_auth.py (Gmail auth)
                   ├── drive_uploader.py (Drive upload)
                   └── xlsx_fixer.py (Excel formatting)
                                ↓
                   report_config.py (all configuration in one place)
```

---

## Setup Instructions

### 1. Prerequisites
- Python 3.13 or higher
- Windows OS (uses win32com for Excel automation)
- Gmail account with API access enabled

### 2. Install Dependencies

Install from requirements.txt:
```bash
pip install -r requirements.txt
```

**Key packages:**
- `google-auth-oauthlib` - Gmail & Drive OAuth authentication
- `google-auth-httplib2` - HTTP library for Google API
- `google-api-python-client` - Gmail & Drive API client
- `pandas` - Data manipulation
- `openpyxl` - Excel file handling
- `pywin32` - Windows COM automation for Excel
- `ttkbootstrap` - Modern themed UI components

### 3. Gmail API Setup

Already configured with `credentials.json`. On first run, the script will:
1. Open a browser window for Gmail authentication
2. Ask you to grant permissions (read, modify, send emails)
3. Save authentication token to `token.pickle` for future runs

### 4. Google Drive Setup (Optional)

For automatic Drive uploads with date-based organization:

1. **Enable Google Drive API** in your Google Cloud Console project
2. **First upload attempt** will open browser for Drive authentication
3. Grant Drive permissions to the app
4. Authentication saves to `token_drive.pickle`

See [DRIVE_SETUP.md](DRIVE_SETUP.md) for detailed configuration.

---

## Configuration

All report configuration lives in **`report_config.py`** - the single source of truth.

### Modifying Recipients

Edit the `ALL_RECIPIENTS` dict and each report's `default_recipients` list:

```python
# report_config.py
ALL_RECIPIENTS = {
    "aidan":    "aidan@tortintakeprofessionals.com",
    "newperson": "newperson@tortintakeprofessionals.com",  # Add new people here
    ...
}

REPORTS = {
    "andy_greg": {
        ...
        "default_recipients": ["aidan", "newperson", ...],  # Reference by short name
    },
}
```

### Adding New Reports

Add subject filters to the appropriate report's `subject_filters` list in `report_config.py`:

```python
REPORTS = {
    "andy_greg": {
        "subject_filters": [
            "Report: A&G: New Report Name - Firm1 - Firm2 - Shield Legal",
            # Add more here...
        ],
    },
}
```

### Drive Folder Configuration

Drive folder IDs are in `drive_uploader.py`:

```python
DRIVE_FOLDERS = {
    "Andy & Greg": "1qOnyoZl_lbWkUGk8r6iMroy8ZTKom91E",
    "Cameron Flatirons": "11V0Ity9HLncxsOkS-yedctUWRO6oiLFF",
    "Cameron & Crump": "1GHYmpl983zq264sZnYj-iR-M_1-w-5LJ",
    "Malissa": "1FroHjovKsopPtRTlr_LPwai7y4isX-DY"
}
```

---

## Workflow Process

### For Each Report Type:

1. **Clear Output Directory** - Deletes existing files, ensures fresh start
2. **Download Reports** - Searches Gmail for matching subjects, downloads .xlsx attachments
3. **Process Each Excel File:**
   - Fixes column widths
   - Converts date columns to MM/DD/YYYY format
   - Renames Sheet1 to "All Fields All Time"
   - Adds formatted Excel table
   - Creates pivot tables with Status field counts
   - Auto-adjusts column widths and reorders sheets
4. **Send Email** (optional) - Attaches all processed .xlsx files
5. **Upload to Drive** (optional) - Automatic date-based folder organization

---

## Report Types

| Report | Filter Count | Pivot Sheets | Config Key |
|--------|-------------|--------------|------------|
| Andy & Greg | 124 | Combined, Matches Dashboard, All Final | `andy_greg` |
| Cameron Flatirons | 56 | Combined, Matches Benchmark, Matches Dashboard | `cameron_flatirons` |
| Cameron & Crump | 3 | Pivot Table | `cameron_crump` |
| Malissa | 22 | Pivot Table | `malissa` |

---

## Common Issues & Solutions

### AttributeError with win32com
```
AttributeError: module 'win32com.gen_py...' has no attribute 'CLSIDToClassMap'
```
**Solution:** Delete the gen_py cache folder:
```powershell
Remove-Item -Recurse "C:\Users\Esteban\AppData\Local\Temp\gen_py"
```

### Gmail API Precondition Failed
**Solution:** Delete token and re-authenticate:
```powershell
Remove-Item token.pickle
# Run script again to re-authenticate
```

### No Emails Found
- No unread emails matching the subject filter
- Reports already processed (emails marked as read)
- Incorrect subject line in `subject_filters` list

---

## Security Notes

- `credentials.json` contains OAuth client secrets
- `token.pickle` / `token_drive.pickle` contain authentication tokens
- **DO NOT** commit these files to public repositories
- All sensitive files are excluded via `.gitignore`

---

## Output Structure

```
Working/Python Outputs/
├── Andy & Greg/
│   ├── Report_1.xlsx
│   └── ...
├── Malissa/
│   └── ...
├── Cameron/
│   ├── Other/
│   │   └── ...
│   └── Flatirons/
│       └── ...
```

Each Excel file contains:
1. "All Fields All Time" sheet (formatted data)
2. Pivot tables (Status counts)

---

## Production Status

**Version:** 2.0 (Refactored)
**Status:** Fully Operational

---

*Last Updated: February 2026*
*Maintained by: Esteban*
