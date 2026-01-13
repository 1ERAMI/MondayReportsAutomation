#Camerons Crump & Other Monday Report Automation

#  %%
import BQSAAuth as BQA
import RepAutoGmail as RAGA  # Import your Gmail authentication module
from Fix_defaultColWidthPt import XLSXFixer
import os
import time
import re
import pandas as pd
import base64
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from base64 import urlsafe_b64decode
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import win32com.client as win32
import shutil
import tempfile

# import win32com.client
# import win32com.client.makepy

# win32com.client.makepy.main()


# win32.gencache.EnsureDispatch('Excel.Application')
# excel = win32.Dispatch('Excel.Application')
gservice = RAGA.confirm_auth()
client = BQA.client

SAVE_DIRECTORY = "C:\\Users\\Aidan\\Desktop\\Working\\Python Outputs\\Cameron\\Other"


def clear_save_directory(directory):
    """
    Clears all files in the specified directory.
    """
    if os.path.exists(directory):
        shutil.rmtree(directory)
        os.makedirs(directory, exist_ok=True)
        print(f"Cleared directory: {directory}")
    else:
        os.makedirs(directory)
        print(f"Created new directory: {directory}")


# %%
# Helper to sanitize filenames
def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "_", filename)

# Function to retrieve email and download the attachment
def get_report_email(gservice, subject_filter):
    query = f"subject:\"{subject_filter}\" is:unread"
    results = gservice.users().messages().list(userId="me", q=query).execute()
    messages = results.get("messages", [])

    if not messages:
        print("No emails found with the given subject.")
        return None

    for message in messages:
        msg = gservice.users().messages().get(userId="me", id=message["id"]).execute()
        payload = msg.get("payload", {})
        parts = payload.get("parts", [])

        for part in parts:
            if part.get("filename") and part.get("mimeType") == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                filename = sanitize_filename(part["filename"])
                body = part.get("body", {})
                if "attachmentId" in body:
                    attachment_id = body["attachmentId"]
                    attachment = gservice.users().messages().attachments().get(
                        userId="me", messageId=message["id"], id=attachment_id
                    ).execute()
                    file_data = base64.urlsafe_b64decode(attachment["data"])

                    # Save the file
                    os.makedirs(SAVE_DIRECTORY, exist_ok=True)
                    file_path = os.path.join(SAVE_DIRECTORY, filename)
                    with open(file_path, "wb") as f:
                        f.write(file_data)
                    print(f"File downloaded: {file_path}")

                    # Mark the email as read
                    gservice.users().messages().modify(
                        userId="me",
                        id=message["id"],
                        body={"removeLabelIds": ["UNREAD"]},
                    ).execute()
                    return file_path

    print("No attachments found.")
    return None

# Process Excel file
def process_excel(file_path):
    """
    Fix Excel issues, clean data, and set date columns to Excel 'Short Date' format.
    """
    try:
        print("Fixing Excel structure...")
        XLSXFixer.fix_default_col_width(file_path)

        print("Processing Excel data...")
        df = pd.read_excel(file_path)

        # Define date columns
        date_columns = ["E-Sign Signed Date", "Lead Created Date", "Date of Birth"]
        
        # Convert columns to datetime
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")

        # Save DataFrame back to Excel
        df.to_excel(file_path, index=False, engine="openpyxl")

        # Now set Excel's native 'Short Date' format (mm/dd/yyyy) using openpyxl
        wb = load_workbook(file_path)
        ws = wb.active
        
        for col_name in date_columns:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name) + 1  # pandas is 0-based, Excel 1-based
                for cell in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    if cell[0].value:
                        cell[0].number_format = 'MM/DD/YYYY'  # Excel short date format

        wb.save(file_path)
        print(f"Dates formatted as 'Short Date'. Final file saved: {file_path}")

        return file_path

    except Exception as e:
        print(f"Error processing the Excel file: {e}")
        return None

# Add Excel table to the first sheet
def add_table_to_sheet(file_path, sheet_name):
    try:
        wb = load_workbook(file_path)
        ws = wb[sheet_name]
        max_row, max_col = ws.max_row, ws.max_column
        table_range = f"A1:{ws.cell(row=max_row, column=max_col).coordinate}"
        table = Table(displayName="Table1", ref=table_range)
        style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True, showColumnStripes=True)
        table.tableStyleInfo = style
        ws.add_table(table)
        wb.save(file_path)
        print(f"Table added to sheet: {sheet_name}")

    except Exception as e:
        print(f"Error adding table to sheet: {e}")

def rename_sheet(file_path, old_name, new_name):
    try:
        wb = load_workbook(file_path)
        if old_name in wb.sheetnames:
            wb[old_name].title = new_name
            wb.save(file_path)
            print(f"Sheet '{old_name}' renamed to '{new_name}'.")
        else:
            print(f"Sheet '{old_name}' not found in the workbook.")
    except Exception as e:
        print(f"Error renaming sheet: {e}")

# Create Excel pivot tables
def create_multiple_pivot_tables(file_path, data_sheet_name, pivot_sheet_names):
    try:
        # clear_gen_py_cache()
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = True  # Show Excel for debugging

        print(f"Opening workbook: {file_path}")
        wb = excel.Workbooks.Open(os.path.abspath(file_path))
        data_sheet = wb.Worksheets(data_sheet_name)
        print(f"Opened sheet: {data_sheet_name}")

        last_row = data_sheet.UsedRange.Rows.Count
        last_col = data_sheet.UsedRange.Columns.Count
        print(f"Detected used range: {last_row} rows, {last_col} columns")

        headers = [data_sheet.Cells(1, col).Value for col in range(1, last_col + 1)]
        print("Headers found in row 1:", headers)

        if "Status" not in headers:
            raise ValueError("The 'Status' column is missing from the data. Cannot create pivot table.")

        data_range = data_sheet.Range(data_sheet.Cells(1, 1), data_sheet.Cells(last_row, last_col))
        pivot_cache = wb.PivotCaches().Create(SourceType=1, SourceData=data_range)

        for pivot_sheet_name in pivot_sheet_names:
            pivot_sheet = wb.Sheets.Add()
            pivot_sheet.Name = pivot_sheet_name
            print(f"Creating pivot table in new sheet: {pivot_sheet_name}")

            pivot_table = pivot_cache.CreatePivotTable(
                TableDestination=pivot_sheet.Cells(1, 1),
                TableName=f"PivotTable_{pivot_sheet_name.replace(' ', '_')}",
            )

            print("Adding fields to pivot table...")

            status_field_row = pivot_table.PivotFields("Status")
            status_field_row.Orientation = 1  # xlRowField
            status_field_row.Position = 1

            status_field_value = pivot_table.PivotFields("Status")
            status_field_value.Orientation = 4  # xlDataField
            status_field_value.Function = -4112  # xlCount
            status_field_value.Name = "Count of Status"

            print(f"Pivot table '{pivot_sheet_name}' created successfully.")

        wb.Worksheets(data_sheet_name).Activate()
        excel.ActiveWindow.View = 1
        excel.ActiveWindow.SelectedSheets(1).Select()

        wb.Save()
        print("All pivot tables created and workbook saved.")

    except Exception as e:
        print(f"Error creating pivot tables: {e}")

    finally:
        if "excel" in locals():
            excel.Quit()

# Format and reorder sheets
def format_and_reorder_sheets(file_path, sheet_order):
    """
    Adjust column widths for the first sheet in the order and reorder sheets.
    Ensures only the first sheet is selected upon saving.
    :param file_path: Path to the Excel file.
    :param sheet_order: List of sheet names in the desired order.
    """
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        print(f"Opening workbook for formatting and reordering: {file_path}")

        wb = excel.Workbooks.Open(os.path.abspath(file_path))

        # Verify sheets exist
        existing_sheets = [sheet.Name for sheet in wb.Sheets]
        for sheet in sheet_order:
            if sheet not in existing_sheets:
                raise ValueError(f"Sheet '{sheet}' not found in workbook.")

        # Adjust column widths for the first sheet
        first_sheet_name = sheet_order[0]
        first_sheet = wb.Worksheets(first_sheet_name)
        first_sheet.Activate()

        print(f"Auto-adjusting column widths for '{first_sheet_name}'...")
        first_sheet.UsedRange.Columns.AutoFit()

        # Reorder sheets explicitly
        print("Reordering sheets...")
        for idx, sheet_name in enumerate(sheet_order, start=1):
            wb.Worksheets(sheet_name).Move(Before=wb.Worksheets(idx))

        # Explicitly select only the first sheet and ungroup sheets
        first_sheet.Select()
        excel.ActiveWindow.View = 1  # xlNormalView ensures ungrouped sheets
        excel.ActiveWindow.SelectedSheets(1).Select()

        wb.Save()
        wb.Close()
        print(f"Sheets reordered and column widths adjusted successfully in '{file_path}'.")

    except PermissionError as e:
        print(f"Permission denied error: {e}")
        print("Ensure the Excel file isn't open elsewhere.")
    except Exception as e:
        print(f"Error during formatting and reordering: {e}")
    finally:
        if "excel" in locals():
            excel.Quit()


def send_email_with_attachments(gservice, to_emails, subject, body, attachment_folder):
    try:
        if isinstance(to_emails, list):
            to_emails = ", ".join(to_emails)

        message = MIMEMultipart()
        message["to"] = to_emails
        message["subject"] = subject
        message.attach(MIMEText(body, "plain"))

        # Attach all .xlsx files in the directory
        for file_name in os.listdir(attachment_folder):
            if file_name.endswith(".xlsx"):
                file_path = os.path.join(attachment_folder, file_name)
                with open(file_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={os.path.basename(file_path)}",
                )
                message.attach(part)

        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        gservice.users().messages().send(userId="me", body={"raw": raw_message}).execute()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Error sending email: {e}")

# def clear_gen_py_cache():
#     """
#     Deletes the 'gen_py' folder in the user's local temp directory to avoid
#     known win32com CLSIDToClassMap errors.
#     """
#     gen_py_path = os.path.join(tempfile.gettempdir(), "gen_py")
#     if os.path.exists(gen_py_path):
#         try:
#             shutil.rmtree(gen_py_path)
#             print(f"Deleted cached COM folder: {gen_py_path}")
#         except Exception as e:
#             print(f"Failed to delete gen_py: {e}")

# Main function
def main():
    """
    Main function to execute the script for multiple reports and send an email with the processed files.
    """

    # Clear directory before processing new reports
    clear_save_directory(SAVE_DIRECTORY)

    # List of subject filters for the reports
    subject_filters = [
        "Report: MAL: Chowchilla Womens Prison Abuse - ACTS - AWD - Shield Legal",
        "Report: MAL: Dr Barry Brock SA - AWD - SGGH - Shield Legal",
        "Report: MAL: Dr Derrick Todd Abuse - SGGH - AWD - Shield Legal",
        "Report: MAL: Dr Scott Lee Abuse - SGGH - AWD - Shield Legal",
        "Report: MAL: Mormon Victim Abuse - Dolman - Anapol Weiss - AWD - Shield Legal",
        "Report: MAL: Mormon Victim Abuse - HRSC - AWD - Shield Legal",
        "Report: MAL: NEC - AWD - Wagstaff - Shield Legal",
        "Report: MAL: Paraquat - AWD - Wagstaff - Shield Legal",
        # # 
        "Report: MAL: Paraquat 2 - AWD - Wagstaff - Shield Legal",
        "Report: MAL: San Bernardino County JDC Abuse - SGGH - AWD - Shield Legal",
        "Report: MAL: San Bernardino County JDC Abuse OSOL - SGGH - AWD - Shield Legal",
        "Report: MAL: San Diego County JDC Abuse - SGGH - AWD - Shield Legal",
        "Report: MAL: San Diego County JDC Abuse OSOL - SGGH - AWD - Shield Legal",
        "Report: MAL: Transvaginal Mesh - Anapol Weiss - AWD - Shield Legal",
        "Report: MAL: Video Gaming Sextortion - Anapol Weiss - AWD - 3PL",
        "Report: MAL: Video Gaming Sextortion - Anapol Weiss - AWD - Shield Legal",
        "Report: MAL: Video Gaming Sextortion - Cooper Masterman - AWD - Shield Legal",
        "Report: MAL: Video Gaming Sextortion SEC - Anapol Weiss - AWD - 3PL",
        "Report: MAL: Video Gaming Sextortion SEC - Anapol Weiss - AWD - Anapol Weiss",
        "Report: MAL: Instant Soup Cup Child Burns - BCL - Gomez - Shield Legal",
        ##
        "Report: MAL: Chowchilla Womens Prison Abuse - Oakwood - Oakwood - Shield Legal",
        "Report: MAL: Polinsky Children's Center Abuse 2 - Oakwood - Oakwood - Shield Legal"
        # Add more subject filters here
    ]

    # Folder to save processed files
    attachment_folder = SAVE_DIRECTORY

    # Email details
    to_emails = [
    "aidan@tortintakeprofessionals.com",
    "ngaston@tortintakeprofessionals.com",
    # "martin@tortintakeprofessionals.com",
    "mclark@tortintakeprofessionals.com",
    # "oroman@tortintakeprofessionals.com",
    "pjerome@tortintakeprofessionals.com"
    # "esteban@tortintakeprofessionals.com",
    # "brittany@tortintakeprofessionals.com",
    # "jackson@tortintakeprofessionals.com"
    ]  # Add recipient emails
    email_subject = "Malissa Monday Reports"
    email_body = (
        "Hello,\n\n"
        "Please find attached the processed reports for Cameron & Crump's Monday Reports.\n\n"
        "Best regards,\n"
        "Your Automation Script ʕ•́ᴥ•̀ʔっ♡"
    )

    # Process each report
    for subject_filter in subject_filters:
        print(f"Processing report for: {subject_filter}")
        file_path = get_report_email(gservice, subject_filter)
        if file_path:
            try:
                # Process the Excel file
                processed_file_path = process_excel(file_path)
                if processed_file_path:
                    rename_sheet(processed_file_path, "Sheet1", "All Fields All Time")
                    add_table_to_sheet(processed_file_path, "All Fields All Time")
                    pivot_sheets = [
                        "Pivot Table"
                        # "Pivot Table Matches Dashboard",
                        # "Pivot Table All Final",
                    ]
                    create_multiple_pivot_tables(processed_file_path, "All Fields All Time", pivot_sheets)
                    format_and_reorder_sheets(
                        processed_file_path,
                        ["All Fields All Time"] + pivot_sheets,
                    )
            except Exception as e:
                print(f"Error processing report: {e}")

        print(f"Finished processing report: {subject_filter}")
        print("-" * 50)

    # Send an email with all processed files
    try:
        print("Sending email with attachments...")
        send_email_with_attachments(gservice, to_emails, email_subject, email_body, attachment_folder)
    except Exception as e:
        print(f"Error sending email: {e}")

    print("All reports processed and email sent successfully.")

if __name__ == "__main__":
    main()


# if you get this error:
# AttributeError: module 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9' has no attribute 'CLSIDToClassMap'
# goto this folder and delete the folder: C:\Users\Aidan\AppData\Local\Temp\gen_py


