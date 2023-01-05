import os
import datetime
import pytz
import calendar
import pandas as pd
import logging

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

LOCAL_TIMEZONE = "Europe/Berlin"

logger = logging.getLogger(__name__)

# The ID of a sample document.
def get_month_datetimes(date=datetime.datetime.today()):
    """Returns the start and end datetime objects for a given month in a datetime object.
    """
    start = date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    end = start + datetime.timedelta(days=calendar.monthrange(start.year, start.month)[1] - 1)
    end = end.replace(hour=23, minute=59, second=59, microsecond=999999)

    return start, end

def convert_datetime_for_api(time:datetime, timezone:str = LOCAL_TIMEZONE):
    """Converts a datetime object to a string in the format required by the Google Calendar API.
    """
    local = pytz.timezone(timezone)
    local_date = local.localize(time)
    utc_date = local_date.astimezone(pytz.utc).isoformat()
    utc_date = str(utc_date).replace('+00:00', 'Z')

    return utc_date

def get_events_by_date(api_service, start_date, end_date, calendar_id, timezone=LOCAL_TIMEZONE):
    """Gets events from a calendar between two dates.
    """
    utc_start_date = convert_datetime_for_api(start_date, timezone=timezone)
    utc_end_date = convert_datetime_for_api(end_date, timezone=timezone)

    logger.info('Searching for events between %s and %s...', utc_start_date, utc_end_date)

    events_result = api_service.events().list(
        calendarId=calendar_id,
        timeMin=utc_start_date,
        timeMax=utc_end_date,
        singleEvents=True
        ).execute()

    logger.info('Found %d events.', len(events_result.get('items', [])))

    return events_result.get('items', [])

def create_events_table(events):
    """Creates a pandas DataFrame from a list of events.
    """
    # API reference:
    # https://developers.google.com/calendar/api/v3/reference/events#resource

    df = pd.DataFrame(columns=['summary', 'start', 'end', 'duration'])

    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        # tz = pytz.timezone(event['start'].get('timeZone')) # Unused
        stop = event['end'].get('dateTime', event['end'].get('date'))

        entry_df = pd.DataFrame({
            'summary': event['summary'],
            'start': pd.to_datetime(start),
            'end': pd.to_datetime(stop),
            'duration': (pd.to_datetime(stop) - pd.to_datetime(start)).total_seconds() / 60 / 60
        }, index=[0], columns=df.columns)

        df = pd.concat([df, entry_df], ignore_index=True)

    return df

def export_stats_by_company(df, export_path = 'export', german=True):
    """Creates a pandas DataFrame with the tota''l work hours per company.
    """

    companies = df['summary'].unique()

    file_paths = []

    for company in companies:
        df_company = df.loc[df['summary'] == company]
        df_company = df_company.reset_index(drop=True)
        month_string = df_company['start'][0].strftime('%Y-%m')
        file_path = export_stats(df_company, f"{export_path}/{company}/{company}_{month_string}.csv", german=german)
        file_paths.append(file_path)

    return file_paths

def export_stats(df, file_path = f"export/{datetime.datetime.today().strftime('%Y-%m')}.csv", german=True):
    """Exports a pandas DataFrame to a csv file.
    """
    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    # Set decimal separator and date format for german or english
    if german:
        sep = ';'
        decimal = ','
        date_format = '%d-%m-%Y %H:%M:%S'
    else:
        sep = ','
        decimal = '.'
        date_format = '%Y-%m-%d %H:%M:%S'
    
    df.to_csv(
        file_path,
        sep=sep,
        decimal=decimal,
        date_format=date_format,
        encoding='utf-8',
        index=False)

    return file_path

def get_event_df_month(creds, month, calendar_id='primary', timezone=LOCAL_TIMEZONE):
    """Returns a pandas DataFrame with all events from a given month.
    """
    try:
        service = build('calendar', 'v3', credentials=creds, cache_discovery=False)

        start_date, end_date = get_month_datetimes(date=month)
        events = get_events_by_date(service, start_date, end_date, calendar_id, timezone=timezone)

        df = create_events_table(events)
    
    except HttpError as error:
        logger.error('Error: %s', error)
        return None

    return df
