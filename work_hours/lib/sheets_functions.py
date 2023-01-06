import pandas as pd
import logging

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logger = logging.getLogger(__name__)

def read_header(creds, spreadsheet_id, sheet_id=0):
    # https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/append
    try:
        # create drive api client
        service = build('sheets', 'v4', credentials=creds, cache_discovery=False)

        # pylint: disable=maybe-no-member
        sheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        
        sheet_0 = sheet['sheets'][sheet_id]['properties']
        sheet_title = sheet_0['title']
        tz = sheet['properties']['timeZone']
        locale = sheet['properties']['locale']

        header_row = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_title}!A1:Z1"
            ).execute()

        if 'values' in header_row:
            header = header_row['values'][0]
        else:
            header = None
        
        logger.info('Got header from Sheet ID: "%s".', spreadsheet_id)

    except HttpError as error:
        logger.error('An error occurred: %s', error)
        header, tz, locale, sheet_title = None, None, None, None

    return header, tz, locale, sheet_title

def append_rows(creds, value_list, spreadsheet_id, sheet='Sheet1'):
    # https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/append
    try:
        # create drive api client
        service = build('sheets', 'v4', credentials=creds, cache_discovery=False)

        range_ = sheet
        value_input_option = 'USER_ENTERED'
        insert_data_option = 'INSERT_ROWS'
        value_range_body = {
            'values': value_list
        }

        # pylint: disable=maybe-no-member
        request = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_,
            valueInputOption=value_input_option,
            insertDataOption=insert_data_option,
            body=value_range_body
            )
        
        response = request.execute()

        logger.info('Sheet with ID: "%s" has been updated.', spreadsheet_id)

    except HttpError as error:
        logger.error('An error occurred: %s', error)
        response = None

    return response

def update_rows(creds, value_list, spreadsheet_id, sheet='Sheet1', start_row=1):
    # https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/append
    try:
        # create drive api client
        service = build('sheets', 'v4', credentials=creds, cache_discovery=False)

        range_ = f"{sheet}!A{start_row}:Z{start_row + len(value_list) - 1}"
        value_input_option = 'USER_ENTERED'
        value_range_body = {
            'values': [value_list]
        }

        # pylint: disable=maybe-no-member
        request = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_,
            valueInputOption=value_input_option,
            body=value_range_body
            )
        
        response = request.execute()

        logger.info('Sheet with ID: "%s" has been updated.', spreadsheet_id)

    except HttpError as error:
        logger.error('An error occurred: %s', error)
        response = None

    return response

def get_statistics_by_company(df):
    df['year'] = df['start'].apply(lambda x: x.strftime('%Y'))
    df['month'] = df['start'].apply(lambda x: x.strftime('%m'))
    
    df = df.groupby(['year', 'month', 'summary'])['duration'].sum()
    df = pd.DataFrame(df).reset_index().pivot(index=['year', 'month'], columns='summary', values='duration').fillna(0)
    # df.index = df.index.astype(str)

    return df

def get_statistics_by_company_weekly(df):
    df['year'] = df['start'].apply(lambda x: pd.to_datetime(x).isocalendar()[0])
    df['week'] = df['start'].apply(lambda x: pd.to_datetime(x).isocalendar()[1])

    df = df.groupby(['year', 'week', 'summary'])['duration'].sum()
    df = pd.DataFrame(df).reset_index().pivot(index=['year', 'week'], columns='summary', values='duration').fillna(0)
    # df.index = df.index.astype(str)

    return df

def sync_header(df, orig_header):
    # Check if there are new columns
    df_columns = [item.lower() for item in df.columns.unique().tolist()]
    orig_header = [item.lower() for item in orig_header]

    new_columns = list(set(df_columns) - set(orig_header))
    header = orig_header + new_columns

    # Get month column
    df_update = df.reset_index().rename(columns={"start": header[0]})

    # All columns to lower case
    df_update.columns = map(str.lower, df_update.columns)
    header = map(str.lower, header)

    # Sort columns
    df_update = df_update.reindex(columns=header)

    # All columns to title case
    df_update.columns = df_update.columns.str.title()

    # Fill NaN with 0 for columns without hours
    df_update = df_update.fillna(0)

    return df_update

def update_sheet(creds, df, spreadsheet_id, time_type='Month'):
    # TODO - Check if row exists

    if time_type == 'Month':
        sheet_id = 0
    elif time_type == 'Week':
        sheet_id = 1
    else:
        raise ValueError('time_type must be Month or Week')
    
    orig_header, _, _, sheet_title = read_header(creds, spreadsheet_id, sheet_id=sheet_id)

    if orig_header is None:
        header = df.columns.tolist()
        header.insert(0, time_type)
        header.insert(0, 'Year')

        append_rows(creds, [header], spreadsheet_id, sheet=sheet_title)

        df_update = sync_header(df, header)
        response = append_rows(creds, df_update.values.tolist(), spreadsheet_id, sheet=sheet_title)
    else:
        df_update = sync_header(df, orig_header)

        if len(df_update.columns) != len(orig_header):
            logger.warning('Columns are not the same. Header: %s, Columns: %s', orig_header, df_update.columns.tolist())
            logger.info('Updating header from Sheet ID: "%s".', spreadsheet_id)
            update_rows(creds, df_update.columns.tolist(), spreadsheet_id, sheet=sheet_title)

        response = append_rows(creds, df_update.values.tolist(), spreadsheet_id, sheet=sheet_title)

    return response


def append_statistics(creds, df, spreadsheet_id, time_type='Month'):
    if time_type == 'Month':
        df = get_statistics_by_company(df)
    else:
        df = get_statistics_by_company_weekly(df)

    logger.info('Summary of work hours for %s: \n %s', time_type.lower(), df.to_string())

    response = update_sheet(creds, df, spreadsheet_id, time_type=time_type)

    return response.get('id')