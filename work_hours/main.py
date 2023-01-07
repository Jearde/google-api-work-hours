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
from oauth2client.service_account import ServiceAccountCredentials

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

parser.add_argument(
    '--week',
    type=int,
    dest='past_week',
    default=0,
    help='Number of months to go back. 0 is current month, 1 is last month (default: 1)')

parser.add_argument(
    '--server',
    type=bool,
    dest='server_mode',
    action=argparse.BooleanOptionalAction,
    help='Use console for authorization (default: False)')

# If modifying these scopes, delete the file token.json.
# https://developers.google.com/identity/protocols/oauth2/scopes#drive
SCOPES = [
    'https://www.googleapis.com/auth/calendar.readonly',
    'https://www.googleapis.com/auth/drive',
    ]

def create_service_credentials(user_email):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        f"{config_path}/service_account.json",
        scopes=SCOPES)

    credentials = credentials.create_delegated(user_email)

    return credentials

def create_token_local(creds, cred_path):
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                f"{cred_path}/credentials.json", SCOPES)
            creds = flow.run_local_server(
                host='localhost',
                port=8088,
                authorization_prompt_message='Please visit this URL for authorizing: {url}',
                success_message='The auth flow is complete; you may close this window.',
                open_browser=True
                )
        # Save the credentials for the next run
        with open(f"{cred_path}/token.json", 'w', encoding='utf-8') as token:
            token.write(creds.to_json())

    return creds


def credentials_to_dict(credentials):
    return {'token': credentials.token,
            'refresh_token': credentials.refresh_token,
            'token_uri': credentials.token_uri,
            'client_id': credentials.client_id,
            'client_secret': credentials.client_secret,
            'scopes': credentials.scopes,
            'id_token': credentials.id_token}

def create_token_server(creds, cred_path):
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            f"{cred_path}/credentials.json",
            scopes=SCOPES
            )

        flow.redirect_uri = "http://localhost"
        auth_url, __ = flow.authorization_url(prompt="consent")
        logger.info('Please go to this URL for authorization: %s', auth_url)
        code = input('Enter the authorization code from the browser URL: ')
        flow.fetch_token(code=code)

        creds = flow.credentials

    # Save the credentials for the next run
    with open(f"{cred_path}/token.json", 'w', encoding='utf-8') as token:
        token.write(creds.to_json())

    return creds

def main(cred_path, config_path, server_mode, month_past=0, week_past=0):
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
        if server_mode:
            creds = create_token_server(creds, cred_path)
        else:
            creds = create_token_local(creds, cred_path)
    
    try:
        with open(f"{config_path}/work_hours.json", encoding='utf-8') as json_file:
            ids = json.load(json_file)

        # Get timezone and locale from Google Sheet
        _, tz, locale, _, sheet_id = gsf.read_header(creds, spreadsheet_id=ids['summary_id'])
        
        # Get month
        today = datetime.datetime.today()
        # first = today.replace(day=1)
        used_month = today + relativedelta(months=-month_past)
        used_week = today + relativedelta(weeks=-week_past)

        # Get events from calendar
        df_week = gcf.get_event_df_week(creds, used_week, calendar_id=ids["calendar_id"], timezone=tz)

        logger.info('Events found week: \n %s', df_week.to_string())

        # Get events from calendar
        df_month = gcf.get_event_df_month(creds, used_month, calendar_id=ids["calendar_id"], timezone=tz)

        logger.info('Events found month: \n %s', df_month.to_string())

        # Delete old csv files
        if os.path.exists(EXPORT_PATH) and os.path.isdir(EXPORT_PATH):
            shutil.rmtree(EXPORT_PATH)

        # Export to csv
        german = locale == 'de_DE'
        file_path = gcf.export_stats(df_month, file_path=f"{EXPORT_PATH}/all_{used_month.strftime('%Y-%m')}.csv", german=german)

        # Export to csv by company
        file_paths = gcf.export_stats_by_company(df_month, export_path=EXPORT_PATH, german=german)

        # Upload all csv files to Google Drive
        gdf.upload_csv_folder_with_conversion(EXPORT_PATH, creds, folder_id=ids["folder_id"])

        # Append monthly sum to Google Sheet
        gsf.append_statistics(creds, df_month, spreadsheet_id=ids['summary_id'], time_type='Month')

        # Append weekly sum to Google Sheet
        gsf.append_statistics(creds, df_week, spreadsheet_id=ids['weekly_id'], time_type='Week')

    except HttpError as error:
        logger.info('An error occurred: %s', error)


if __name__ == '__main__':
    args = parser.parse_args()

    log_path = os.path.abspath(args.log_path)
    cred_path = os.path.abspath(args.cred_path)
    config_path = os.path.abspath(args.config_path)
    past_month = args.past_month
    past_week = args.past_week
    server_mode = args.server_mode

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

    main(cred_path, config_path, server_mode, month_past=past_month, week_past=past_week)
