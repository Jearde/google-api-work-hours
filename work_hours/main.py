import argparse
import logging
import os.path
import shutil
import datetime
from dateutil.relativedelta import relativedelta
import json

import lib.calendar_functions as gcf
import lib.drive_functions as gdf
import lib.sheets_functions as gsf

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError

EXPORT_PATH = 'export'

logger = logging.getLogger(__name__)

parser = argparse.ArgumentParser(description='Google Calendar Work Hours')
parser.add_argument(
    '--log',
    dest='log_path',
    default='logs',
    help='Path to log directory (default: logs)')

parser.add_argument(
    '--cred',
    type=str,
    dest='cred_path',
    default='config',
    help='Path to directory with credentials.json (default: config)')

parser.add_argument(
    '--config',
    type=str,
    dest='config_path',
    default='config',
    help='Path to directory with work_hours.json (default: config)')

parser.add_argument(
    '--month',
    type=int,
    dest='past_month',
    default=0,
    help='Number of months to go back. 0 is current month, 1 is last month (default: 1)')

# If modifying these scopes, delete the file token.json.
# https://developers.google.com/identity/protocols/oauth2/scopes#drive
SCOPES = [
    'https://www.googleapis.com/auth/calendar.readonly',
    'https://www.googleapis.com/auth/drive',
    ]    

def main(month_past, cred_path, config_path):
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(f"{cred_path}/token.json"):
        creds = Credentials.from_authorized_user_file(f"{cred_path}/token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                f"{cred_path}/credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(f"{cred_path}/token.json", 'w', encoding='utf-8') as token:
            token.write(creds.to_json())
    
    try:
        with open(f"{config_path}/work_hours.json", encoding='utf-8') as json_file:
            ids = json.load(json_file)

        # Get timezone and locale from Google Sheet
        _, tz, locale, _ = gsf.read_header(creds, spreadsheet_id=ids['summary_id'])
        
        # Get month
        today = datetime.datetime.today()
        first = today.replace(day=1)
        used_month = first + relativedelta(months=-month_past)

        # Get events from calendar
        df = gcf.get_event_df_month(creds, used_month, calendar_id=ids["calendar_id"], timezone=tz)

        logger.info('Events found: \n %s', df.to_string())

        # Delete old csv files
        if os.path.exists(EXPORT_PATH) and os.path.isdir(EXPORT_PATH):
            shutil.rmtree(EXPORT_PATH)

        # Export to csv
        german = locale == 'de_DE'
        file_path = gcf.export_stats(df, file_path=f"{EXPORT_PATH}/all_{used_month.strftime('%Y-%m')}.csv", german=german)

        # Export to csv by company
        file_paths = gcf.export_stats_by_company(df, export_path=EXPORT_PATH, german=german)

        # Upload all csv files to Google Drive
        gdf.upload_csv_folder_with_conversion(EXPORT_PATH, creds, folder_id=ids["folder_id"])

        # Append monthly sum to Google Sheet
        file_id = gsf.append_statistics(creds, df, spreadsheet_id=ids['summary_id'])

    except HttpError as error:
        logger.info('An error occurred: %s', error)


if __name__ == '__main__':
    args = parser.parse_args()

    log_path = os.path.abspath(args.log_path)
    cred_path = os.path.abspath(args.cred_path)
    config_path = os.path.abspath(args.config_path)
    past_month = args.past_month

    os.makedirs(log_path, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%d-%m-%y %H:%M:%S',
        handlers=[
            logging.FileHandler(f"{log_path}/work_hours.log"),
            logging.StreamHandler()
        ]
    )

    main(past_month, cred_path, config_path)
