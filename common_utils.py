# Importing required modules
import pandas as pd
import win32com.client
import os
import time
import pythoncom
from datetime import datetime, timezone, timedelta

import logging, logging.config
import yaml
from datetime import datetime, timedelta

"""
Basic utility class with collection of some of the 
repetitive functions being used through out the application. 
"""

with open("logging.yml", "r") as file:
    config = yaml.safe_load(file)
    logging.config.dictConfig(config)

logger = logging.getLogger('my_logger')


def connect_outlook(profile_email):
    """Connect to Outlook and selesct profile with the given mail ID"""
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Get available mails
    available_profiles = [profile.Name for profile in outlook.Folders]
    
    if profile_email not in available_profiles:
        raise ValueError(f"Profile '{profile_email}' not found.\nAvailble profiles: {available_profiles}")
                         
    return outlook.Folders(profile_email)


def get_inbox_messages(folder, start_date, end_date):
    logger.info(f'Start Date: {start_date}\nEnd Date: {end_date}')
    try:
        start_date = datetime.strptime(start_date,"%d-%m-%Y").strftime("%d-%m-%Y 00:00 AM")
        # end_date = (start_date + timedelta(days=1) - timedelta(seconds=1))
        end_date = datetime.strptime(end_date,"%d-%m-%Y").strftime("%d-%m-%Y 11:59 PM")
        inbox = folder.Folders("Inbox")
        messages = inbox.Items

        if start_date and end_date:
            messages = messages.Restrict("[ReceivedTime]>='{}' AND [ReceivedTime]<'{}'"
                                            .format(start_date,
                                                    end_date))
            logger.info(f"Fetched {messages.Count} tips emails for dates between {start_date} and {end_date}")
        return messages
    except Exception as e:
        logger.error(f"Error in fetching emails: {e}")


def extract_latest_mail(mail_body):
    """Extract only the latest response from an email before the reply chain"""
    reply_seperators=[
        "From:", "Sent:", "To:", "Subject:"
    ]
    
    # Split by new lines
    lines = mail_body.split("\n")
    
    # Find where the reply chain starts
    for i, line in enumerate(lines):
        if any(seperator in line for seperator in reply_seperators):
            return "\n".join(lines[:i]).strip()
        
    return mail_body.strip()


def extract_mail_details(message_):
    """Extract mail details into a dictionary"""
    
    subject = getattr(message_, "Subject", "")
    classification = "C2/C3"
    try:
        sender = getattr(message_, "SenderName", "")
        to_recipients = getattr(message_, "To", "")
        cc_recipient = getattr(message_, "CC", "") 
        bcc_recipient = getattr(message_, "BCC", "") 
        senttime = getattr(message_, "SentOn", None) 
        if senttime is not None:
            senttime = str(senttime)
            senttime = datetime.fromisoformat(senttime).replace(microsecond = 0).strftime("%Y-%m-%d %H:%M:%S") 
        received_time= getattr(message_, "ReceivedTime", None)
        if received_time is not None:
            received_time = str(received_time)
            received_time = datetime.fromisoformat(received_time).replace(microsecond = 0).strftime("%Y-%m-%d %H:%M:%S")
    
        body = getattr(message_, "Body", "") 
        latest_body = extract_latest_mail(body)

        importance = "Low"
        keywords = ["importance: high", "EOD", "immediately", "Urgent", "as soon as possible", "by Today", "by tomorrow", "by next week"]
        today = datetime.now()
        date_keywords = [
            today.strftime("%d-%m-%Y"),
            (today + timedelta(days=1)).strftime("%d-%m-%Y"),
            (today + timedelta(weeks=1)).strftime("%d-%m-%Y"),
        ]

        # Check for keywords or dates near today in the email body
        if any(keyword.lower() in body.lower() for keyword in keywords) or any(date in body for date in date_keywords):
            importance = "High"
            
        suspicious_email = "No"
        if "[EMETTEUR EXTERNE] / [EXTERNAL SENDER]" in body.upper():
            suspicious_email = "Yes"
            
        attactments = []
        if suspicious_email == "No":
            attactments = [att.FileName for att in message_.Attachments]

        attactments_present = "Yes" if attactments else "No"
        if suspicious_email == "Yes":
            attactments_present = ""
        
        classification = "Not C2/C3"
        return {
                "Subject":subject,
                "Sender": sender,
                "To" : to_recipients,
                "CC": cc_recipient,
                "BCC" : bcc_recipient,
                "Sent Time": senttime,
                "Received Time" : received_time,
                "Mail Content": body,
                "Latest Mail Content": latest_body,
                "Attachments Present(Yes/No)" : attactments_present,
                "Attachment Names": ", ".join(dict.fromkeys(attactments)) if attactments else "",
                "Classification": classification,
                "Importance": importance,
                "Suspicious": suspicious_email
            }
    except:
        return {
                "Subject":subject,
                "Sender": "",
                "To" : "",
                "CC": "",
                "BCC" : "",
                "Sent Time": "",
                "Received Time" : "",
                "Mail Content": "",
                "Latest Mail Content": "",
                "Attachments Present(Yes/No)" : "",
                "Attachment Names": "",
                "Classification": classification,
                "Importance": "NA",
                "Suspicious": "NA"
            }
    

def download_attachments(message, sender_name, save_path):
    """Download all attachments from an email"""
    logger.info(f"Downloading attachments for email from {sender_name}")
    try:
        if not os.path.exists(save_path):
            os.makedirs(save_path)
        
        for att in message.Attachments:
            att.SaveAsFile(os.path.join(save_path, f"{sender_name}_{att.FileName}"))
        logger.info(f"Attachments saved to {save_path}")
    except:
        logger.error(f"Error in downloading attachments for email from {sender_name}")


def process_emails(profile_name, start_date, end_date, save_path="attachements"):
    """Main function to process emails and return a DataFrame"""
    outlook_folder = connect_outlook(profile_name)
    messages = get_inbox_messages(outlook_folder, start_date, end_date)
    email_data = []
    for message in messages:
        try:
            email_details = extract_mail_details(message)
            logger.info("Email details extracted")
            email_data.append(email_details)
            # message.display()

            # Download attachment if present
            if email_details.get('Attachments Present(Yes/No)', "")== "Yes":
                if email_details.get('Suspicious', "")== "No":
                    sender_name = email_details['Sender']
                    download_attachments(message, sender_name, save_path)
                
            time.sleep(3)
            message.Close(0)
        except Exception as e:
            logger.error(f'Could not display message: {e}')   
        
    email_details_df = pd.DataFrame(email_data)
    return email_details_df