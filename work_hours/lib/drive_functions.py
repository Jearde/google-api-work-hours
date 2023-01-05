import os
import logging

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

logger = logging.getLogger(__name__)

def check_if_file_exists(creds, file_name, folder_id):
    """Check if a file with the same name already exists in the folder.
    """

    # API reference: 
    # https://developers.google.com/drive/api/v3/reference/files/list

    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds, cache_discovery=False)

        # Check if a file with the same name already exists in the folder
        query = f"'{folder_id}' in parents and trashed = false and name='{file_name}'"
        
        # pylint: disable=maybe-no-member
        files = service.files().list(q=query, fields='nextPageToken, '
                                                    'files(id, name)').execute()
        existing_files = files.get('files', [])
       
        if existing_files:
            # A file with the same name already exists, so return the existing file
            logger.info('File with name "%s" already exists.', file_name)
            return existing_files[0]
        else:
            return False

    except HttpError as error:
        logger.error('An error occurred: %s', error)
        return None

def upload_csv_with_conversion(file_path, creds, folder_id):
    """Upload a csv file to Google Drive and convert it to a Google Sheet.
    """

    # API reference: 
    # https://developers.google.com/drive/api/guides/manage-uploads

    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds, cache_discovery=False)

        base = os.path.basename(file_path)
        file_name = os.path.splitext(base)[0]

        file_response = check_if_file_exists(creds, file_name, folder_id)

        if file_response:
            logger.warning('Skipping file with name "%s" and ID "%s"', file_response.get("name"), file_response.get("id"))
            return file_response.get('id')

        # create file metadata
        file_metadata = {
            'name': file_name,
            'parents': [folder_id],
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        }

        # create media body
        media = MediaFileUpload(file_path, mimetype='text/csv',
                                resumable=True)

        # pylint: disable=maybe-no-member
        request = service.files().create(body=file_metadata, media_body=media,
                                      fields='id, name')
        response = request.execute()

        logger.info('File "%s" with ID: "%s" has been uploaded.', response.get("name"), response.get("id"))

    except HttpError as error:
        logger.error('An error occurred: %s', error)
        response = None

    return response.get('id')

def create_folder(creds, folder_name, parent_folder_id):
    """Create a folder in Google Drive.
    """

    # API reference: 
    # https://developers.google.com/drive/api/v3/reference/files/create

    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds, cache_discovery=False)

        # Check if a folder with the same name already exists in the parent folder
        query = f"'{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed = false and name='{folder_name}'"
        
        # pylint: disable=maybe-no-member
        files = service.files().list(q=query, fields='nextPageToken, '
                                                    'files(id, name)').execute()
        existing_folders = files.get('files', [])
       
        if existing_folders:
            # A folder with the same name already exists, so return the existing folder
            logger.info('Folder with name "%s" already exists.', folder_name)
            return existing_folders[0]
        else:
            # Create the folder metadata
            mimetype = 'application/vnd.google-apps.folder'
            file_metadata = {'name': folder_name, 'mimeType': mimetype}
            if parent_folder_id:
                file_metadata['parents'] = [parent_folder_id]

            # Create the folder
            # pylint: disable=maybe-no-member
            file_response = service.files().create(body=file_metadata, fields='id,name').execute()
            logger.info('Folder has been created with Name "%s" and URL: "https://drive.google.com/drive/folders/%s".', folder_name, file_response.get("id"))

        return file_response
    
    except HttpError as error:
        logger.error('An error occurred while creating the folder: %s', error)
        file_response = None
    
    return file_response

def upload_csv_folder_with_conversion(export_path, creds, folder_id):
    """Upload a csv file to Google Drive and convert it to a Google Sheet.
    """

    # API reference: 
    # https://developers.google.com/drive/api/guides/manage-uploads

    for directory in os.listdir(export_path): # Get all the directories in the current working directory
        directory_path = os.path.join(export_path, directory)
        if os.path.isdir(directory_path): # Check if the directory is a folder
            g_folder = create_folder(creds, directory, folder_id) # Create a folder in Google Drive
            for file in os.listdir(directory_path): # Get all the files in the folder
                file_path = os.path.join(directory_path, file) # Get the file path
                if os.path.isfile(file_path): # Check if the file is a file
                    if file.endswith('.csv'): # Check if the file is a csv file
                        upload_csv_with_conversion(file_path, creds, g_folder['id']) # Upload the csv file to Google Drive