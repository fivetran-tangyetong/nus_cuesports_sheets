import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

import datetime

import Constants

def auth() -> Credentials:
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(Constants.TOKEN):
        creds = Credentials.from_authorized_user_file(Constants.TOKEN, Constants.SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(Constants.CREDENTIALS, Constants.SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(Constants.TOKEN, 'w') as token:
            token.write(creds.to_json())
    
    return creds


# day: 0 = Monday, 1 = Tuesday, 2 = Wednesday...
def next_weekday(day):
    currDate = datetime.date.today()
    days = (day - currDate.weekday() + 7) % 7
    return datetime.datetime.strptime(str(currDate + datetime.timedelta(days=days)), "%Y-%m-%d")


def processExcelChr(idx) -> str:
    if idx + Constants.MASTER_DATE_START + 65 > 90:
        return 'A' + chr(idx + Constants.MASTER_DATE_START + 39)
    
    return chr(idx + Constants.MASTER_DATE_START + 65)


def parseColumnByDate(sheet, nextDate) -> (str, str):
    dates = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=Constants.MASTER_DATE_RANGE).execute()
    dates_values = dates.get('values', [])[0]

    if not dates_values:
            print('No data found!' + str(nextDate))
            quit()

    for (idx, date) in enumerate(dates_values):
        if datetime.datetime.strptime(date, "%d/%m/%Y") == nextDate:
            # + 4 because started from 4th index onwards (Counting from E)
            # + 65 to return capital letter
            # This only works for F - AZ (BA onwards will not work)
            return processExcelChr(idx)
        
    print("Can't find next Date! " + str(nextDate))
    return ''

def parseColumnByDates(sheet, nextMon) -> (str, str):
    dates = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=Constants.MASTER_DATE_RANGE).execute()
    dates_values = dates.get('values', [])[0]

    if not dates_values:
            print('No data found!' + str(nextMon))
            quit()

    for (idx, date) in enumerate(dates_values):
        if datetime.datetime.strptime(date, "%d/%m/%Y") == nextMon:
            # + 4 because started from 4th index onwards (Counting from E)
            # + 65 to return capital letter
            # This only works for F - AZ (BA onwards will not work)
            return (processExcelChr(idx), processExcelChr(idx + 1))
        
    print("Can't find next Monday! " + str(nextMon))
    quit()