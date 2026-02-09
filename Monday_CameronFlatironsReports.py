import RepAutoGmail as RAGA  # Import your Gmail authentication module
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
from Fix_defaultColWidthPt import XLSXFixer
import shutil

# win32.gencache.EnsureDispatch('Excel.Application')
excel = win32.Dispatch('Excel.Application')
gservice = RAGA.confirm_auth()

SAVE_DIRECTORY = "C:\\Users\\Esteban\\Desktop\\Working\\Python Outputs\\Cameron\\Flatirons"

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

# Process Excel file with pandas
def process_excel(file_path):
    """
    Fix Excel issues, clean data, and save the final version using the original filename.
    """
    try:
        print("Fixing Excel structure...")
        XLSXFixer.fix_default_col_width(file_path)

        print("Processing Excel data...")
        df = pd.read_excel(file_path)

        # Format date columns
        date_columns = ["E-Sign Signed Date", "Lead Created Date", "Date of Birth"]
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")  # Keep as datetime objects

        # Save the cleaned data back to the Excel file
        df.to_excel(file_path, index=False, engine="openpyxl")

        # Format date columns in Excel using openpyxl
        wb = load_workbook(file_path)
        ws = wb.active
        for col_name in date_columns:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name) + 1
                for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    cell = row[0]
                    if cell.value:
                        cell.number_format = 'MM/DD/YYYY'  # Excel 'Short Date' format

        wb.save(file_path)
        print(f"Final file saved: {file_path}")
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
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(os.path.abspath(file_path))
        data_sheet = wb.Worksheets(data_sheet_name)

        last_row = data_sheet.UsedRange.Rows.Count
        last_col = data_sheet.UsedRange.Columns.Count
        data_range = data_sheet.Range(data_sheet.Cells(1, 1), data_sheet.Cells(last_row, last_col))

        pivot_cache = wb.PivotCaches().Create(SourceType=1, SourceData=data_range)

        for pivot_sheet_name in pivot_sheet_names:
            pivot_sheet = wb.Sheets.Add()
            pivot_sheet.Name = pivot_sheet_name
            pivot_table = pivot_cache.CreatePivotTable(
                TableDestination=pivot_sheet.Cells(1, 1),
                TableName=f"PivotTable_{pivot_sheet_name.replace(' ', '_')}",
            )
            status_field_row = pivot_table.PivotFields("Status")
            status_field_row.Orientation = 1
            status_field_row.Position = 1
            status_field_value = pivot_table.PivotFields("Status")
            status_field_value.Orientation = 4
            status_field_value.Function = -4112  # xlCount
            status_field_value.Name = "Count of Status"

        # Explicitly select only one sheet (deselect grouped sheets)
        wb.Worksheets(data_sheet_name).Activate()
        excel.ActiveWindow.View = 1  # xlNormalView
        excel.ActiveWindow.SelectedSheets(1).Select()

        wb.Save()
        print("Pivot tables created successfully.")

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



# Main function
def main():
    """
    Main function to execute the script for multiple reports and send an email with the processed files.
    """

    # Clear directory before processing new reports
    clear_save_directory(SAVE_DIRECTORY)

    # List of subject filters for the reports
    subject_filters = [
        "Report: MFI: AFFF-PFAS Military Base Exposure - DL - Flatirons - Shield Legal",
        "Report: MFI: Alameda County Juv Hall - DL - Flatirons - Shield Legal",
        "Report: MFI: AZ YRTC - DL - Flatirons - Shield Legal",
        "Report: MFI: Bard PowerPort - DL - Flatirons - Shield Legal",
        "Report: MFI: CA Juv Hall Abuse - ACTS - DL Flatirons - Shield Legal",
        "Report: MFI: Camp Lejeune - DL - Flatirons - Shield Legal",
        "Report: MFI: Camp Lejeune EO - DL - Flatirons - Shield Legal",
        "Report: MFI: Chowchilla Womens Prison Abuse - ACTS - DiCello/Flatirons (TIP) - Shield Legal",
        "Report: MFI: Chowchilla Womens Prison Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: CPAP - DL - Flatirons - Shield Legal",
        "Report: MFI: Depo-Provera - DL - Flatirons - Shield Legal",
        "Report: MFI: Ethylene Oxide - DL - Flatirons - Shield Legal",
        "Report: MFI: Firefighting Foam - DL - Flatirons - Shield Legal",
        "Report: MFI: Hair Relaxer - DL - Flatirons - Shield Legal",
        "Report: MFI: Hair Relaxer Cancer - DL - BCL - DeMayo Flatirons - BLX",
        "Report: MFI: Hair Relaxer Cancer - DL - BCL - DeMayo Flatirons - Shield Legal",
        "Report: MFI: Hair Relaxer Cancer - DL - DeMayo Flatirons - Shield Legal",
        "Report: MFI: Hair Salon Bladder Cancer - DL - Flatirons - Shield Legal",
        "Report: MFI: Illinois Juv Hall Abuse TV - DL/BG - Flatirons - Shield Legal",
        "Report: MFI: Illinois Juvenile Hall Abuse - Dicello - BG/Flatirons NC (TIP) - Shield Legal",
        "Report: MFI: Illinois Juvenile Hall Abuse - Dicello - Flatirons NC (TIP) - Shield Legal",
        "Report: MFI: Instant Soup Cup Child Burns - DL - Flatirons - Shield Legal",
        "Report: MFI: Instant Soup Cup Child Burns 2 - DL - Flatirons - Shield Legal",
        "Report: MFI: Kanakuk Kamps Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: LA County Foster Care Abuse - ACTS - DL/Flatirons - Shield Legal",
        "Report: MFI: LA Juvenile Hall Abuse TV - ACTS - DL/Flatirons - Shield Legal",
        "Report: MFI: LA Wildfires - DL - Flatirons - Shield Legal",
        "Report: MFI: LA Wildfires Flyer - DL - Flatirons - Shield Legal",
        "Report: MFI: Los Padrinos Juv Hall Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: MD Juv Abuse TV - DiCello Levitt - DL/Flatirons - Shield Legal",
        "Report: MFI: MD Juv Hall Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: MD Juv Hall Abuse - DL/BG - Flatirons - Shield Legal",
        "Report: MFI: Michigan Juv Hall Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: Michigan Juv Hall Abuse - DL/BG - Flatirons - Shield Legal",
        "Report: MFI: Mormon Victim Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: Mormon Victim Abuse OSOL - DL - Flatirons - Shield Legal",
        "Report: MFI: NEC Baby Formula - DL - Flatirons - Shield Legal",
        "Report: MFI: New Hampshire YDC Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: NJ Juvenile Hall Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: NV YRTC - DL - Flatirons - Shield Legal",
        "Report: MFI: PA Juvenile Hall Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: Paraquat 1 - DL - DeMayo Flatirons - Shield Legal",
        "Report: MFI: Paraquat 12 - DL - Flatirons - Shield Legal",
        "Report: MFI: Paraquat 5 - DL - DeMayo Flatirons - Shield Legal",
        "Report: MFI: Paraquat 7 - DL - DeMayo Flatirons - Shield Legal",
        "Report: MFI: Polinsky Childrens Center Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: Porterville Developmental Center Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: Private Boarding School Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: Riverside JDC Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: Sacramento JDC Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: San Bernardino Juv Hall - DL - Flatirons - Shield Legal",
        "Report: MFI: San Diego Juv Hall YRTC - DL - Flatirons - Shield Legal",
        "Report: MFI: Santa Clara County Juv Hall - DL - Flatirons - Shield Legal",
        "Report: MFI: Tepezza - DL - Flatirons - Shield Legal",
        "Report: MFI: Trinity Private School Abuse - DL - Flatirons - Shield Legal",
        "Report: MFI: NEC Baby Formula 2 - DL - Flatirons - Shield Legal"


        # Add more subject filters here
    ]

    # Folder to save processed files
    attachment_folder = SAVE_DIRECTORY

    # Email details
    to_emails = [#"aidan@tortintakeprofessionals.com",
                #  "martin@tortintakeprofessionals.com",
                #"ngaston@tortintakeprofessionals.com",
                #  "oroman@tortintakeprofessionals.com", 
                "esteban@tortintakeprofessionals.com" 
                #  "brittany@tortintakeprofessionals.com", 
                #  "jackson@tortintakeprofessionals.com"
                 ]  # Add recipient emails
    email_subject = "Cameron Flatirons Reports"
    email_body = (
        "Hello,\n\n"
        "Please find attached the processed reports for Camerons Flatirons Reports.\n\n"
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
                        "Pivot Table Combined",
                        "Pivot Table Matches Benchmark",
                        "Pivot Table Matches Dashboard",
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