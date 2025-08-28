# Importing required modules
import pandas as pd
import win32com.client
import os
import time
import pythoncom

from datetime import datetime, timezone, timedelta
from common_utils import *

def run(start_date, end_date):
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        available_profiles = [profile.Name for profile in outlook.Folders]
        logger.info(f"Available profiles: {available_profiles}")
    except Exception as e:
        logger.info(f"Error initializing Outlook: {e}")

    logger.info(f"Start Date: {start_date}\nEnd Date: {end_date}")

    profile_name = "saranya.govindaraj@socgen.com"
    # date = input("Enter the date for which you want to read the email in 'DD-MM-YYYY': ")

    # start_date = '23-07-2025'
    # end_date = '23-07-2025' 

    savepath = r"C:\Users\sgovinda021323\OneDrive - GROUP DIGITAL WORKPLACE\Documents\Hackathon_2025\email_extraction\attachments"
    fin_email_details_df = process_emails(profile_name, start_date, end_date, savepath)
    logger.info(fin_email_details_df.shape)
    fin_email_details_df.to_csv("csv_files/email_details.csv")
    return fin_email_details_df
