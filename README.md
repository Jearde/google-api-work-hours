# google-api-work-hours
A Python Programm using the Google Workspace API for exporting work hours from calendar to a Google Sheet.

## Run manually

### Install dependencies

```bash
python3.9 -m venv env
source env/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
``` 

## Google Drive Preparation
1. Create a Google Drive folder for the statistics
    1. Add the ID of the folder to the `work_hours.json` file under folder_id
2. Create a Google Sheet in the folder
    1. Name the first sheet "Sheet1"
    2. (Optional) Add the needed columns in the first sheet: `"Month", <Company_1>, <Company_1>, ..., <Company_n>`
    3. Add the ID of the sheet to the `work_hours.json` file under sheet_id
3. Create a Google Calendar for the work hours
    1. Add the ID of the calendar to the `work_hours.json` file under calendar_id
    2. Add events for the time worked to the calendar with the following format: `<Company>`. Example: `Company A`


## What it does
1. Gets all calendar entries from the calendar with specified ID
2. Saves the calendar entries in a Google Sheet with specified folder
3. Creates a new folder for each name found in the calendar entries (company names)
4. Creates a new sheet for each name found in the calendar entries (hours per month for each company)
5. Appends the hours per month to a specified main sheet