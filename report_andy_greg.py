"""
Andy & Greg Monday Report Automation
Thin wrapper around report_common.run_report() with 'andy_greg' config.
"""

from report_common import run_report


def main(to_emails=None, status_callback=None, send_email=True, upload_to_drive=False):
    """Entry point called by MondayReportsUI.py or run standalone."""
    run_report(
        "andy_greg",
        to_emails=to_emails,
        status_callback=status_callback,
        send_email=send_email,
        upload_to_drive=upload_to_drive,
    )


if __name__ == "__main__":
    main()
