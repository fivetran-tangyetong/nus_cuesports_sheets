from __future__ import print_function

import os.path
import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a master tracking spreadsheet.
SPREADSHEET_ID = '1Wvk4UGKgiT_UACsgzUSepydVA7GgO5FbqtsL9p3qM-4'

# [[Date]]
MEMBER_TRACKING_SHEET_DATE_RANGE = 'Sem 1!E2:AF2'
# [[Full Name, Telegram Display name, Telegram Handle]]
MEMBER_TRACKING_SHEET_NAME_RANGE = 'Sem 1!A19:C'

# [[Date]]
MEMBER_ATTENDANCE_SHEET_DATE_RANGE = 'Aug 23!F8:K8'
# [[Full Name, Telegram Handle]]
MEMBER_ATTENDANCE_SHEET_NAME_RANGE = 'Aug 23!B10:C121'

TOKEN = 'token.json'
CREDENTIALS = 'credentials.json'

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
    if os.path.exists(TOKEN):
        creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(TOKEN, 'w') as token:
            token.write(creds.to_json())
    
    return creds


def parseColumnByDates(sheet, nextMon) -> (str, str):
    dates = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=MEMBER_TRACKING_SHEET_DATE_RANGE).execute()
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


def getNamesOfComing(sheet, COMING_RANGE) -> list:
    coming = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=COMING_RANGE).execute()
    coming_values = coming.get('values', [])

    if not coming_values:
        print('No data found for those who PAID!')
        quit()

    coming_list = []
    for (idx, paid_value) in enumerate(coming_values):
        if len(paid_value) != 0:
            coming_list.append(idx)
    
    return coming_list


def filterComing(sheet, coming_list) -> list:
    names = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=MEMBER_TRACKING_SHEET_NAME_RANGE).execute()
    names_values = names.get('values', [])

    if not names_values:
        print('Tracking Sheet Name Empty!')
        quit()

    filtered_list = []
    for (idx, names_value) in enumerate(names_values):
        if idx in coming_list:
            filtered_list.append(names_value)

    return filtered_list


def checkAttendanceSheet(sheet, day, coming_list) -> None:
    trainingDates = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=MEMBER_ATTENDANCE_SHEET_DATE_RANGE).execute()
    trainingDates_values = trainingDates.get('values', [])[0]

    if not trainingDates_values:
        print('Attendance Sheet Date!')
        quit()

    dateIdx = chr(trainingDates_values.index(day) + 5 + 65)

    currAttendance = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=MEMBER_ATTENDANCE_SHEET_NAME_RANGE).execute()
    currAttendance_values = currAttendance.get('values', [])

    if not currAttendance_values:
        print('Attendance Sheet Name!')
        quit()
        
    currAttendance_telegram_values = []
    currAttendance_telegram_values.append([i[1] for i in currAttendance_values])
    currAttendance_telegram_values = currAttendance_telegram_values[0]

    coming_telegram_values = []
    coming_telegram_values.append([i[2] for i in coming_list])
    coming_telegram_values = coming_telegram_values[0]

    currAttendanceSize = len(currAttendance_telegram_values)

    

    for coming_telegram_value in coming_telegram_values:
        # Add their name and telegram handle if not in attendance sheet
        if coming_telegram_value not in currAttendance_telegram_values:
            currAttendanceSize += 1
            currAttendanceName = ""
            for coming_list_value in coming_list:
                if coming_telegram_value in coming_list_value:
                    currAttendanceName = coming_list_value[0]
            
            new_attendance_range = 'Aug 23!A' + str(currAttendanceSize + 9)
            new_attendance_row = [currAttendanceSize, currAttendanceName, coming_list_value]

            new_attendance_request = {
                "values": [new_attendance_row],
            }

            sheet.values().append(
                spreadsheetId=SPREADSHEET_ID,
                range=new_attendance_range,
                body=new_attendance_request,
                valueInputOption="RAW"
            ).exeute()

            print("Added " + coming_telegram_value + " at row " + str(coming_attendance_row))
        
        coming_attendance_row = currAttendance_telegram_values.index(coming_telegram_value) + 1 + 9
        coming_attendance_range = 'Aug 23!' + dateIdx + str(coming_attendance_row)
        coming_attendance_value = '0'

        coming_attendance_update_values_request = {
            "range": coming_attendance_range,
            "majorDimension": "ROWS",
            "values": [[coming_attendance_value]],
        }

        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=coming_attendance_range,
            body=coming_attendance_update_values_request,
            valueInputOption="RAW"
        ).execute()

        print("Updated " + coming_telegram_value + " at row " + str(coming_attendance_row))

def main() -> None:
    creds = auth()

    nextMon = next_weekday(0)

    try:
        service = build('sheets', 'v4', credentials=creds)
        # Call the Sheets API
        sheet = service.spreadsheets()

        (monIdx, wedIdx) = parseColumnByDates(sheet, nextMon)

        MEMBER_TRACKING_SHEET_COMING_MONDAY_RANGE = 'Sem 1!' + monIdx + '19:' + monIdx
        MEMBER_TRACKING_SHEET_COMING_WEDNESDAY_RANGE = 'Sem 1!' + wedIdx + '19:' + wedIdx

        monNames = getNamesOfComing(sheet, MEMBER_TRACKING_SHEET_COMING_MONDAY_RANGE)
        wedNames = getNamesOfComing(sheet, MEMBER_TRACKING_SHEET_COMING_WEDNESDAY_RANGE)

        monNames_filtered = filterComing(sheet, monNames)
        wedNames_filtered = filterComing(sheet, wedNames)

        checkAttendanceSheet(sheet, str(nextMon.day), monNames_filtered)
        checkAttendanceSheet(sheet, str(nextMon.day + 2), wedNames_filtered)

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()