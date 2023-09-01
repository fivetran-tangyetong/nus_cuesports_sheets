from __future__ import print_function

import os.path
import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import Constants

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


def getMembersPaid(sheet) -> (list, list):
    data = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_FORM, range=Constants.FORM_DATA_RANGE).execute()
    data_values = data.get('values', [])

    if not data_values:
            print('No data found!')
            quit()

    paid_list = []

    for data_value in data_values:
        if data_value[Constants.FORM_TELE_ID][0] != '@':
            data_value[Constants.FORM_TELE_ID] = '@' + data_value[Constants.FORM_TELE_ID]

        person_name = data_value[Constants.FORM_NAME_ID].strip()
        person_year = ''
        person_tele = data_value[Constants.FORM_TELE_ID].strip().lower()
        person_date = data_value[Constants.FORM_DATE_ID]

        if ' ' in data_value[Constants.FORM_YEAR_ID].strip():
            person_year = data_value[Constants.FORM_YEAR_ID].strip().split(" ")[1]
        else:
            person_year = data_value[Constants.FORM_YEAR_ID].strip()

        paid_list.append([person_name, person_year, person_tele, person_date])

    return paid_list


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
            return (chr(idx + 4 + 65), chr(idx + 1 + 4 + 65))
        
    print("Can't find next Monday! " + str(nextMon))
    quit()


def updateMembersPaid(sheet, paid_list, mon_idx, wed_idx) -> None:
    print("Updating Members Paid!")
    names = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=Constants.MASTER_NAME_RANGE).execute()
    names_values = names.get('values', [])

    if not names_values:
        print('Tracking Sheet Name Empty!')
        quit()

    telegram_values = [i[2].lower() for i in names_values]
    update_list = []

    for paid_value in paid_list:
        if paid_value[Constants.FORM_TELE_ID] in telegram_values:

            date_idx = []
            if paid_value[Constants.FORM_DATE_ID] == Constants.MON:
                date_idx = [mon_idx]
            elif paid_value[Constants.FORM_DATE_ID] == Constants.WED:
                date_idx = [wed_idx]
            else:
                date_idx = [mon_idx, wed_idx]

            for date in date_idx:
                member_idx_raw = telegram_values.index(paid_value[Constants.FORM_TELE_ID])
                member_idx = str(member_idx_raw + int(Constants.MASTER_DATA_START)) 
                member_date_range = Constants.MASTER_SHEET_NAME + "!" + date + member_idx + ":" + date + member_idx
                update_list.append({"range": member_date_range, "values": [[Constants.PAID]]})

                if len(names_values[member_idx_raw][Constants.FORM_NAME_ID]) == 0:
                    member_name_range = Constants.MASTER_SHEET_NAME + "!" + Constants.MASTER_ID_NAME + member_idx + ":" + Constants.MASTER_ID_NAME + member_idx
                    update_list.append({"range": member_name_range, "values": [[paid_value[Constants.FORM_NAME_ID]]]})

                if len(names_values[member_idx_raw][Constants.FORM_YEAR_ID]) == 0:
                    member_year_range = Constants.MASTER_SHEET_NAME + "!" + Constants.MASTER_ID_YEAR + member_idx + ":" + Constants.MASTER_ID_YEAR + member_idx
                    update_list.append({"range": member_year_range, "values": [[paid_value[Constants.FORM_YEAR_ID]]]})

        else:
            print(paid_value[Constants.FORM_TELE_ID] + " not in list!")
            
    update_values_request = {
        "valueInputOption": "USER_ENTERED",
        "data": update_list,
    }

    sheet.values().batchUpdate(
        spreadsheetId=Constants.SPREADSHEET_ID_MASTER,
        body=update_values_request,
    ).execute()


def main() -> None:
    creds = auth()

    nextMon = next_weekday(0)

    try:
        service = build('sheets', 'v4', credentials=creds)
        # Call the Sheets API
        sheet = service.spreadsheets()

        paid_list = getMembersPaid(sheet)
        (monIdx, wedIdx) = parseColumnByDates(sheet, nextMon)
        
        updateMembersPaid(sheet, paid_list, monIdx, wedIdx)

    except HttpError as err:
        print(err)
        
    print("Finished! Cleaning up...")


if __name__ == '__main__':
    main()