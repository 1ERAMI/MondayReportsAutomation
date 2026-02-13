"""
Centralized configuration for all Monday report automations.
Edit this file to add/remove reports, change recipients, or modify settings.
"""

import os

# Base output directory (all reports save under this)
BASE_OUTPUT_DIR = os.path.join(
    os.path.expanduser("~"), "Desktop", "Working", "Python Outputs"
)

# All known email recipients (defined once, referenced by short name)
ALL_RECIPIENTS = {
    "aidan":    "aidan@tortintakeprofessionals.com",
    "martin":   "martin@tortintakeprofessionals.com",
    "ngaston":  "ngaston@tortintakeprofessionals.com",
    "oroman":   "oroman@tortintakeprofessionals.com",
    "pjerome":  "pjerome@tortintakeprofessionals.com",
    "esteban":  "esteban@tortintakeprofessionals.com",
    "brittany": "brittany@tortintakeprofessionals.com",
    "jackson":  "jackson@tortintakeprofessionals.com",
    "mclark":   "mclark@tortintakeprofessionals.com",
}

# Per-report configurations
REPORTS = {
    "andy_greg": {
        "display_name": "Andy & Greg Reports",
        "save_subdir": "Andy & Greg",
        "drive_folder_name": "Andy & Greg",
        "email_subject": "Andy & Greg's Monday Reports",
        "email_body": (
            "Hello,\n\n"
            "Please find attached the processed reports for Andy & Greg's Monday Reports.\n\n"
            "Best regards,\n"
            "Your Automation Script \u0295\u2022\u0301\u1d25\u2022\u0300\u0294\u3063\u2661"
        ),
        "default_recipients": [
            "aidan", "martin", "ngaston", "pjerome",
            "esteban", "brittany", "jackson", "mclark",
        ],
        "pivot_sheets": [
            "Pivot Table Combined",
            "Pivot Table Matches Dashboard",
            "Pivot Table All Final",
        ],
        "subject_filters": [
            "Report: A&G: Bard-PowerPort-Bay-Point-Simmons-Shield-Legal",
            "Report: A&G: CA-Juvenile-Hall-Abuse-Miller-Mattar-Shield-Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse - ACTS - AWD - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse-ACTS-Banafshe-Shield Legal",
            "Report: A&G: Chowchilla-ACTS/DL-Flatirons-Shield-Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse - ACTS - Ghozland - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse - AWKO - Goldlaw - Shield Legal",
            "Report: A&G: Cumberland-Hospital-Abuse-BB-BB-Shield Legal",
            "Report: A&G: Dr Barry Brock SA - AWD - SGGH - Shield Legal",
            "Report: A&G: Dr Derrick Todd Abuse - SGGH - AWD - Shield Legal",
            "Report: A&G: Dr Scott Lee Abuse - SGGH - AWD - Shield Legal",
            "Report: A&G: Firefighting Foam - BG - Ghozland - Shield Legal",
            "Report: A&G: Firefighting Foam - Douglas London - ML - Shield Legal",
            "Report: A&G: Firefighting Foam - ELG - ML - Shield Legal",
            "Report: A&G: Firefighting Foam - ELG - SNL - Shield Legal",
            "Report: A&G: Firefighting Foam - ELG - Ye - Shield Legal",
            "Report: A&G: Firefighting-Foam-Napoli-Shkolnik-AFK",
            "Report: A&G: Fire-Fighting-Foam-Nations-Levinson-TC-Shield-Legal",
            "Report: A&G: Firefighting-Foam-Nations-Van-Shield Legal",
            "Report: A&G: Firefighting-Foam-VAM-Law-ELG-Meadow-Law",
            "Report: A&G: Firefighting Foam 2 - VAM Law - ELG - Meadow Law - Shield Legal",
            "Report: A&G: Gaming Addiction 2 - AWKO - Bradley Grombacher - Shield Legal",
            "Report: A&G: Gaming Addiction 3 - AWKO - Bradley Grombacher - Shield Legal",
            "Report: A&G: MI Juv Hall Abuse - Ghozland - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: IL Clergy Abuse - SGGH - GLF - Shield Legal",
            "Report: A&G: IL YTC Abuse - SGGH - GLF - Shield Legal",
            "Report: A&G: Illinois-Juv-Hall-Abuse-Bailey-Glasser-Van-Shield-Legal",
            "Report: A&G: Illinois-Juv-Hall-Abuse-BG-DLA-Shield-Legal",
            "Report: A&G: Illinois Juvenile Hall Abuse - BG - Bay Point - Shield Legal",
            "Report: A&G: MD Juv Hall-BG-FLG-Shield Legal",
            "Report: A&G: MD Juv Hall - BG - Forman - Shield Legal",
            "Report: A&G: MD-Juv-Hall-BG-Ghozland-Shield-Legal",
            "Report: A&G: MD-Juv-Hall-Abuse-Bailey-Glasser-Van-Shield-Legal",
            "Report: A&G: MD-Juvenile-Hall-Abuse-Bailey-Glasser-Rhine-Shield-Legal",
            "Report: A&G: MD Juvenile Hall Abuse - BG - Meadow Law - Shield Legal",
            "Report: A&G: MI Clergy Abuse - SGGH - GLF - Shield Legal",
            "Report: A&G: MI YTC Abuse - SGGH - GLF - Shield Legal",
            "Report: A&G: Mormon Victim Abuse - Dolman - Anapol Weiss - AWD - Shield Legal",
            "Report: A&G: Mormon Victim Abuse - HRSC - AWD - Shield Legal",
            "Report: A&G: New Hampshire YDC Abuse - BG - Ghozland - Shield Legal",
            "Report: A&G: New Hampshire YDC Abuse - BG - VAM Law - Meadow Law - Shield Legal",
            "Report: A&G: New Hampshire YDC Abuse - BG - Van - Shield Legal",
            "Report: A&G: PA-Juv-Hall-BG-Bowersox-Shield-Legal",
            "Report: A&G: PA-Juv-Hall-Abuse-BG-Ghozland-Shield-Legal",
            "Report: A&G: PA-Juv-Hall-Abuse-Lakin-AFG-Levinson-Shield Legal",
            "Report: A&G: PA-Juvenile-Hall-Abuse-Lakin-Lakin-Shield-Legal",
            "Report: A&G: Paraquat-AWD-Wagstaff-Shield Legal",
            "Report: A&G: Paraquat - DL - ML - Shield Legal",
            "Report: A&G: Paraquat - LegaFi - DL - Shield Legal",
            "Report: A&G: Paraquat - LegaFi - Wagstaff - Shield Legal",
            "Report: A&G: Paraquat - Wagstaff - ML - Shield Legal",
            "Report: A&G: PFAS-Napoli-Shkolnik-AFK-Shield-Legal",
            "Report: A&G: PFAS-Water-Contamination-Baypoint-ELG-Shield-Legal",
            "Report: A&G: San Diego County JDC Abuse - SGGH - AWD - Shield Legal",
            "Report: A&G: San Bernardino County JDC Abuse - SGGH - AWD - Shield Legal",
            "Report: A&G: Transvaginal Mesh - Anapol Weiss - AWD - Shield Legal",
            "Report: A&G: Video-Gaming-Sextortion-JB-JB-Shield Legal",
            "Report: A&G: Depo-Provera - Meadow Law - Seeger Weiss - Shield Legal",
            "Report: A&G: Illinois YRTC Abuse - Bay Point - SGGH - Shield Legal",
            "Report: A&G: MI Clergy Abuse - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: MI Foster Care Abuse - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: MI Juv Hall Abuse - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: MI YRTC Abuse - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: MI YRTC Abuse OSOL - AFG - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: Mormon Victim Abuse - AWKO - AWKO - Shield Legal",
            "Report: A&G: Video Gaming Sextortion - Ghozland - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: NEC Baby Formula 2 - LegaFi - LegaFi - Shield Legal",
            "Report: A&G: Mormon Victim Abuse OSOL - WS - WS - Shield Legal",
            "Report: A&G: Mormon Victim Abuse OSOL - AWKO - AWKO - Shield Legal",
            "Report: A&G: MI YRTC Abuse - AFG - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: KS YRTC Abuse - Meadow Law - AWKO - Shield Legal",
            "Report: A&G: KS JDC Abuse - Meadow Law - AWKO - Shield Legal",
            "Report: A&G: Depo Provera - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse - Meadow - SGGH - Shield Legal",
            "Report: A&G: MI Juv Hall Abuse OSOL - AFG - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: MI YRTC Abuse - Ghozland - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: MI YRTC Abuse OSOL - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: MI YRTC Abuse - OSOL - SGGH - Meadow Law - Shield Legal",
            "Report: A&G: Mormon Victim Abuse - Marsh - HRSC - Shield Legal",
            "Report: A&G: Mormon Victim Abuse - WS - WS - Shield Legal",
            "Report: A&G: Northwell Sleep Center - LCA - Slater - Alliant - Shield Legal",
            "Report: A&G: Video Gaming Sextortion - Wright & Schulte - Wright & Schulte - Shield Legal",
            "Report: A&G: MI Clergy Abuse - AFG - SGGH - Meadow - Shield Legal",
            "Report: A&G: MI Clergy Abuse - SGGH - Meadow - Ghozland - Shield Legal",
            "Report: A&G: MI Clergy Abuse - SGGH - Meadow - Yih - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse - SGGH - Milberg - Sanders - Shield Legal",
            "Report: A&G: AZ YTC Abuse - Burg Simpson - ML - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse 2 - SGGH - Meadow - Shield Legal",
            "Report: A&G: Dr Oumair Aejaz Abuse - SGGH - SGGH - Shield Legal",
            "Report: A&G: NEC Baby Formula 2 - AWKO - AWKO - Shield Legal",
            "Report: A&G: Backpage Remission/Hotel/Technology - Ghozland - Ghozland",
            "Report: A&G: Backpage Remission/Hotel/Technology - Rochen - Rochen",
            "Report: A&G: Alameda JDC Abuse - Pulaski - Stinar Lannen",
            "Report: A&G: CA JDC Abuse - Bay Point - GGH - Shield Legal",
            "Report: A&G: CA JDC Abuse - Pulaski - Stinar Lannen - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison - Meadow - Meadow - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse - Abelson - Pulaski - Stinar Lannen - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse - GED - DL - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse - Meadow - Fong - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse - Pulaski - Stinar Lannen - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse - SGGH - Meadow - Baypoint - Shield Legal",
            "Report: A&G: Dupixent CTCL Leukemia - AWKO - AWKO - Shield Legal",
            "Report: A&G: IL JDC Abuse - WS - UT - Shield Legal",
            "Report: A&G: Illinois Juvenile Hall Abuse - WS - WS - Shield Legal",
            "Report: A&G: Instant Soup Cup Child Burns - Meadow - Gomez - Shield Legal",
            "Report: A&G: Instant Soup Cup Child Burns - Meadow - Lanier - Shield Legal",
            "Report: A&G: Mormon Victim Abuse - UT - WS - Shield Legal",
            "Report: A&G: Mormon Victim Abuse OSOL - UT - WS - Shield Legal",
            "Report: A&G: NEC Baby Formula - AWKO - Ghozland - Shield Legal",
            "Report: A&G: NEC Baby Formula - TorHoerman - Meadow - Shield Legal",
            "Report: A&G: NEC Baby Formula 3 - AWKO - AWKO - Shield Legal",
            "Report: A&G: Nevada YRTC Abuse - AWKO - AWKO - Shield Legal",
            "Report: A&G: New Jersey JDC Abuse - AWKO - AWKO - Shield Legal",
            "Report: A&G: Pennsylvania JDC Abuse - AWKO - AWKO - Shield Legal",
            "Report: A&G: Pennsylvania JDC Abuse - Cochran - Flanagan - Shield Legal",
            "Report: A&G: Pennsylvania Juvenile Sexual Abuse - Constant - Constant - Shield Legal",
            "Report: A&G: Riverside JDC Abuse - Pulaski - Stinar Lannen - Shield Legal",
            "Report: A&G: San Bernardino County JDC Abuse - Bay Point - GGH - Shield Legal",
            "Report: A&G: San Bernardino JDC Abuse - Pulaski - Stinar Lannen - Shield Legal",
            "Report: A&G: Video Gaming Sextortion - Bradley Grombacher - Bradley Grombacher - Shield Legal",
            "Report: A&G: Video Gaming Sextortion - Meadow - Meadow - Shield Legal",
            "Report: A&G: Chowchilla Womens Prison Abuse - AFG - Meadow - Shield Legal",
        ],
    },

    "cameron_flatirons": {
        "display_name": "Cameron Flatirons Reports",
        "save_subdir": os.path.join("Cameron", "Flatirons"),
        "drive_folder_name": "Cameron Flatirons",
        "email_subject": "Cameron Flatirons Reports",
        "email_body": (
            "Hello,\n\n"
            "Please find attached the processed reports for Camerons Flatirons Reports.\n\n"
            "Best regards,\n"
            "Your Automation Script \u0295\u2022\u0301\u1d25\u2022\u0300\u0294\u3063\u2661"
        ),
        "default_recipients": [
            "aidan", "martin", "ngaston", "oroman",
            "esteban", "brittany", "jackson",
        ],
        "pivot_sheets": [
            "Pivot Table Combined",
            "Pivot Table Matches Benchmark",
            "Pivot Table Matches Dashboard",
        ],
        "subject_filters": [
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
            "Report: MFI: NEC Baby Formula 2 - DL - Flatirons - Shield Legal",
            "Report:MFI: CA JDC Abuse - DL - Flatirons - Shield Legal"
        ],
    },

    "cameron_crump": {
        "display_name": "Cameron & Crump Reports",
        "save_subdir": os.path.join("Cameron", "Other"),
        "drive_folder_name": "Cameron & Crump",
        "email_subject": "Cameron & Crump's Reports",
        "email_body": (
            "Hello,\n\n"
            "Please find attached the processed reports for Cameron & Crump's Monday Reports.\n\n"
            "Best regards,\n"
            "Your Automation Script \u0295\u2022\u0301\u1d25\u2022\u0300\u0294\u3063\u2661"
        ),
        "default_recipients": [
            "aidan", "martin", "oroman", "pjerome",
            "esteban", "brittany", "jackson",
        ],
        "pivot_sheets": [
            "Pivot Table",
        ],
        "subject_filters": [
            "Report: BCL: Chowchilla Womens Prison Abuse - ACTS - Crump",
            "Report: BCL: Illinois Juvenile Hall Abuse - Crump - Slater",
            "Report: CAM: Nursing Home & Assisted Living Abuse - MRW - MRW",
        ],
    },

    "malissa": {
        "display_name": "Malissa Reports",
        "save_subdir": "Malissa",
        "drive_folder_name": "Malissa",
        "email_subject": "Malissa Monday Reports",
        "email_body": (
            "Hello,\n\n"
            "Please find attached the processed reports for Malissa's Monday Reports.\n\n"
            "Best regards,\n"
            "Your Automation Script \u0295\u2022\u0301\u1d25\u2022\u0300\u0294\u3063\u2661"
        ),
        "default_recipients": [
            "aidan", "ngaston", "martin", "mclark", "oroman",
            "pjerome", "esteban", "brittany", "jackson",
        ],
        "pivot_sheets": [
            "Pivot Table",
        ],
        "subject_filters": [
            "Report: MAL: Chowchilla Womens Prison Abuse - ACTS - AWD - Shield Legal",
            "Report: MAL: Dr Barry Brock SA - AWD - SGGH - Shield Legal",
            "Report: MAL: Dr Derrick Todd Abuse - SGGH - AWD - Shield Legal",
            "Report: MAL: Dr Scott Lee Abuse - SGGH - AWD - Shield Legal",
            "Report: MAL: Mormon Victim Abuse - Dolman - Anapol Weiss - AWD - Shield Legal",
            "Report: MAL: Mormon Victim Abuse - HRSC - AWD - Shield Legal",
            "Report: MAL: NEC - AWD - Wagstaff - Shield Legal",
            "Report: MAL: Paraquat - AWD - Wagstaff - Shield Legal",
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
            "Report: MAL: Chowchilla Womens Prison Abuse - Oakwood - Oakwood - Shield Legal",
            "Report: MAL: Polinsky Children's Center Abuse 2 - Oakwood - Oakwood - Shield Legal",
        ],
    },
}


def get_save_directory(report_key):
    """Get the full save directory path for a report."""
    return os.path.join(BASE_OUTPUT_DIR, REPORTS[report_key]["save_subdir"])


def get_default_emails(report_key):
    """Resolve shorthand recipient names to full email addresses."""
    return [ALL_RECIPIENTS[name] for name in REPORTS[report_key]["default_recipients"]]
