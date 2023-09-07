from __future__ import print_function

import os.path
import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import Constants

spreadsheet_dates = []
spreadsheet_names = []

# day: 0 = Monday, 1 = Tuesday, 2 = Wednesday...
def next_weekday(day):
    currDate = datetime.date.today()
    days = (day - currDate.weekday() + 7) % 7
    return datetime.datetime.strptime(str(currDate + datetime.timedelta(days=days)), "%Y-%m-%d")


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


def getComing(file) -> list:
    contents = []
    with open(file, 'r') as f:
        for line in f.readlines():
            if ',' in line:
                line_parsed = line.split(",")
                contents.append([line_parsed[0].strip(), line_parsed[1].strip()])
            else:
                contents.append([line.strip(), Constants.COMING])        
        f.close()
    
    return contents

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
            return (chr(idx + Constants.MASTER_DATE_START + 65), chr(idx + 1 + Constants.MASTER_DATE_START + 65))
        
    print("Can't find next Monday! " + str(nextMon))
    quit()


def updateMembers(sheet, coming_list, date_idx) -> None:
    print("Updating Members Coming!")
    names = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=Constants.MASTER_NAME_RANGE).execute()
    names_values = names.get('values', [])
    names_len = len(names_values)

    if not names_values:
        print('Tracking Sheet Name Empty!')
        quit()

    telegram_values = [i[Constants.MASTER_ID_TELE_LIST] for i in names_values]
    update_list = []

    for coming_value in coming_list:
        if coming_value[0] in telegram_values:
            member_idx = str(telegram_values.index(coming_value[0]) + int(Constants.MASTER_DATA_START))
            member_date_range = Constants.MASTER_SHEET_NAME + "!" + date_idx + member_idx + ":" + date_idx + member_idx
            update_list.append({"range": member_date_range, "values": [[coming_value[1]]]})
        else:
            member_idx = str(names_len + int(Constants.MASTER_DATA_START))
            member_tele_range = Constants.MASTER_SHEET_NAME + "!" + Constants.MASTER_ID_TELE + member_idx + ":" + Constants.MASTER_ID_TELE + member_idx
            member_date_range = Constants.MASTER_SHEET_NAME + "!" + date_idx + member_idx + ":" + date_idx + member_idx
            update_list.append({"range": member_tele_range, "values": [[coming_value[0]]]})
            update_list.append({"range": member_date_range, "values": [[coming_value[1]]]})

            names_len += 1
            
    update_values_request = {
        "valueInputOption": "USER_ENTERED",
        "data": update_list,
    }

    sheet.values().batchUpdate(
        spreadsheetId=Constants.SPREADSHEET_ID_MASTER,
        body=update_values_request,
    ).execute()


def checkAllMembers(sheet, coming_list, COMING_RANGE) -> None:
    print("Checking All Members if Coming!")
    coming = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=COMING_RANGE).execute()
    coming_values = coming.get('values', [])
    
    names = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=Constants.MASTER_NAME_RANGE).execute()
    names_values = names.get('values', [])

    if not coming_values:
        print('No data found for those who PAID!')
        quit()
        
    if not names_values:
        print('Tracking Sheet Name Empty!')
        quit()

    telegram_values = [i[Constants.MASTER_ID_TELE_LIST] for i in names_values]
    coming_list_names = [i[0] for i in coming_list]

    for (idx, coming_value) in enumerate(coming_values):
        if len(coming_value) != 0 and telegram_values[idx] not in coming_list_names:
            print(telegram_values[idx] + " is COMING, but not in list!")


def main() -> None:
    creds = auth()

    nextMon = next_weekday(0)

    coming_mon = getComing(Constants.TELE_MON)
    coming_wed = getComing(Constants.TELE_WED)

    try:
        service = build('sheets', 'v4', credentials=creds)
        # Call the Sheets API
        sheet = service.spreadsheets()

        (monIdx, wedIdx) = parseColumnByDates(sheet, nextMon)
        
        updateMembers(sheet, coming_mon, monIdx)
        updateMembers(sheet, coming_wed, wedIdx)

        MEMBER_TRACKING_SHEET_COMING_MONDAY_RANGE = Constants.MASTER_SHEET_NAME + '!' + monIdx + Constants.MASTER_DATA_START + ':' + monIdx
        MEMBER_TRACKING_SHEET_COMING_WEDNESDAY_RANGE = Constants.MASTER_SHEET_NAME + '!' + wedIdx + Constants.MASTER_DATA_START + ':' + wedIdx

        checkAllMembers(sheet, coming_mon, MEMBER_TRACKING_SHEET_COMING_MONDAY_RANGE)
        checkAllMembers(sheet, coming_wed, MEMBER_TRACKING_SHEET_COMING_WEDNESDAY_RANGE)        

    except HttpError as err:
        print(err)
        
    print("Finished! Cleaning up...")


if __name__ == '__main__':
    main()