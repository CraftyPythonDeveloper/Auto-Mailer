__author__ = "Amit Yadav"
__email__ = "amityadav4664@gmail.com"

import os
import sys
import logging
import warnings
from datetime import datetime
from collections.abc import Iterable
import pandas as pd
from win32com.client.gencache import EnsureDispatch

# all paths
WRK_DIR = os.path.dirname(os.path.realpath(__file__))
os.chdir(WRK_DIR)
CONFIG_PATH = os.path.join(WRK_DIR, "config.xlsx")
log_file_path = os.path.join(WRK_DIR, "auto_mailer.logs")

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
file_handler = logging.FileHandler(log_file_path)
file_formatter = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s')
file_handler.setFormatter(file_formatter)
console_handler = logging.StreamHandler()
console_formatter = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s')
console_handler.setFormatter(console_formatter)
logger.addHandler(file_handler)
logger.addHandler(console_handler)


def get_outlook_app():
    try:
        # Generate the Outlook.Application class
        outlook = EnsureDispatch("Outlook.Application")
    except Exception as e:
        print(f"Error: {e}")
        outlook = None
    return outlook


def get_df_value(df, lookup_column, lookup_value, column, multi=False):
    try:
        if multi:
            return df.loc[df[lookup_column] == lookup_value, column].values
        return df.loc[df[lookup_column] == lookup_value, column].values[0]
    except IndexError:
        return None


def send_outlook_email(subject, body, to_addresses, attachment_path=None):
    outlook_app = get_outlook_app()
    mail = outlook_app.CreateItem(0)
    mail.Subject = subject
    mail.HTMLBody = body
    if isinstance(to_addresses, Iterable) and not isinstance(to_addresses, str):
        mail.To = ";".join(to_addresses)
    else:
        mail.To = to_addresses
    if isinstance(attachment_path, str):
        mail.Attachments.Add(attachment_path)
    mail.Send()
    logger.info(f"Email sent successfully at {datetime.now()}")
    return True


def get_draft_email_html(email_subject):
    outlook_app = get_outlook_app()
    namespace = outlook_app.GetNamespace("MAPI")
    drafts_folder = namespace.GetDefaultFolder(16)
    for item in drafts_folder.Items:
        if item.Subject == email_subject:
            return item.HTMLBody
    logger.error("No Matching email subject found in outlook draft. \n"
                 "Make sure you have a email subject matching to excel saved in draft")
    input("\nPress any key to exit..")
    return None


if __name__ == "__main__":
    if len(sys.argv) != 2:
        logger.error("Invalid command-line arguments. Usage: python auto_mailer.py <group_id>")
        sys.exit(1)

    group_id = sys.argv[1]

    if " " in group_id:
        logger.info("group_id should not contain any spaces and should be unique in message sheet.")
        input("Press any key to exit")
        sys.exit()

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=UserWarning)
        email_df = pd.read_excel(CONFIG_PATH, sheet_name="emails")
        message_df = pd.read_excel(CONFIG_PATH, sheet_name="messages")

    email_df.group_id = email_df.group_id.astype(str)
    message_df.group_id = message_df.group_id.astype(str)

    if group_id not in email_df.group_id.tolist() or group_id not in message_df.group_id.tolist():
        logger.error(f"Group ID {group_id} not found in excel..")
        input("Press any key to exit..")
        sys.exit()

    subject = get_df_value(message_df, "group_id", group_id, "subject")
    message = get_draft_email_html(subject)
    attachment = get_df_value(message_df, "group_id", group_id, "attachment")
    emails = get_df_value(email_df, "group_id", group_id, "email", True)
    try:
        status = send_outlook_email(subject=subject, body=message, to_addresses=emails.tolist(),
                                    attachment_path=attachment)
    except Exception as e:
        logger.error(f"Error sending email: {str(e)}")
    logger.info("Script execution completed.")
