import pandas as pd
import logging

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logger = logging.getLogger(__name__)

def read_header(creds, spreadsheet_id, sheet_idx=0):
    # https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/append
    try:
        # create drive api client
        service = build('sheets', 'v4', credentials=creds, cache_discovery=False)

        # pylint: disable=maybe-no-member
        sheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        
        sheet_0 = sheet['sheets'][sheet_idx]['properties']
        sheet_title = sheet_0['title']
        sheet_id = sheet_0['sheetId']
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

    return header, tz, locale, sheet_title, sheet_id

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

def update_spreadsheet(cred, spreadsheet_id, data, sheet='Sheet1', sheet_id=0):
    try:
        # Create a service client and build the Sheets API service
        service = build('sheets', 'v4', credentials=cred)

        # Get the data from the sheet
        # pylint: disable=maybe-no-member
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=f'{sheet}!A:Z').execute()
        values = result.get('values', [])

        # String values to float for comparison
        values_num = [[float(v.replace(',','.')) if v != '' else 0 for v in val] for val in values[1:]]
        data_num = [[float(v) for v in val] for val in data]

        # Create a dictionary of rows to update and rows to append
        rows_to_update = {}
        rows_to_append = []
        for row in data_num:
            found = False
            for i, r in enumerate(values_num):
                if len(r) == 0: # Skip empty rows in sheet
                    continue
                if r[0] == row[0] and r[1] == row[1]:
                    rows_to_update[i+1] = row
                    found = True
                    # break # Stop if the first match is found
            if not found:
                rows_to_append.append(row)

        # Update the rows in the sheet
        requests = []
        for i, row in rows_to_update.items():
            requests.append({
                'updateCells': {
                    'start': {
                        'sheetId': sheet_id,
                        'rowIndex': i,
                        'columnIndex': 0
                    },
                'rows': [{
                    'values': [{'userEnteredValue': {'numberValue': cell}} for cell in row]
                }],
                'fields': 'userEnteredValue',
                }
            })
        if requests:
            # pylint: disable=maybe-no-member
            response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={'requests': requests}).execute()

        # Append the rows to the sheet
        if rows_to_append:
            # pylint: disable=maybe-no-member
            response = service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id, range=f'{sheet}!A:Z',
                insertDataOption='INSERT_ROWS', valueInputOption='USER_ENTERED',
                body={'values': rows_to_append}).execute()

        logger.info('Spreadsheet with ID "%s"  and sheet %d has been updated.', spreadsheet_id, sheet_id)
        logger.info('Updated rows: %s', rows_to_update)
        logger.info('Appended rows: %s', rows_to_append)

    except HttpError as error:
        logger.error('An error occurred while updating the spreadsheet: %s', error)
        response = None


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
    if time_type == 'Month':
        sheet_id = 0
    elif time_type == 'Week':
        sheet_id = 1
    else:
        raise ValueError('time_type must be Month or Week')
    
    orig_header, _, _, sheet_title, sheet_id = read_header(creds, spreadsheet_id, sheet_idx=sheet_id)

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

        update_spreadsheet(creds, spreadsheet_id, df_update.values.tolist(), sheet=sheet_title, sheet_id=sheet_id)
        # response = append_rows(creds, df_update.values.tolist(), spreadsheet_id, sheet=sheet_title)


def append_statistics(creds, df, spreadsheet_id, time_type='Month'):
    if time_type == 'Month':
        df = get_statistics_by_company(df)
    else:
        df = get_statistics_by_company_weekly(df)

    logger.info('Summary of work hours for %s: \n %s', time_type.lower(), df.to_string())

    update_sheet(creds, df, spreadsheet_id, time_type=time_type)