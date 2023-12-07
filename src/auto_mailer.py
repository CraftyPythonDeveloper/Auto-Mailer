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
    mail.Body = body
    if isinstance(to_addresses, Iterable):
        mail.To = ";".join(to_addresses)
    else:
        mail.To = to_addresses
    if isinstance(attachment_path, str):
        mail.Attachments.Add(attachment_path)
    mail.Send()
    logging.info(f"Email sent successfully at {datetime.now()}")
    return True


if __name__ == "__main__":
    if len(sys.argv) != 2:
        logging.error("Invalid command-line arguments. Usage: python auto_mailer.py <group_id>")
        sys.exit(1)

    group_id = sys.argv[1]

    if " " in group_id:
        logging.info("group_id should not contain any spaces and should be unique in message sheet.")
        input("Press any key to exit")
        sys.exit()

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=UserWarning)
        email_df = pd.read_excel(CONFIG_PATH, sheet_name="emails")
        message_df = pd.read_excel(CONFIG_PATH, sheet_name="messages")

    email_df.group_id = email_df.group_id.astype(str)
    message_df.group_id = message_df.group_id.astype(str)

    if group_id not in email_df.group_id.tolist() or group_id not in message_df.group_id.tolist():
        print(f"Group ID {group_id} not found in excel..")
        input()
        sys.exit()

    subject = get_df_value(message_df, "group_id", group_id, "subject")
    message = get_df_value(message_df, "group_id", group_id, "message")
    attachment = get_df_value(message_df, "group_id", group_id, "attachment")
    emails = get_df_value(email_df, "group_id", group_id, "email", True)
    try:
        status = send_outlook_email(subject=subject, body=message, to_addresses=emails.tolist(),
                                    attachment_path=attachment)
    except Exception as e:
        logging.error(f"Error sending email: {str(e)}")

    logging.info("Script execution completed.")
