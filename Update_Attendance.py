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


def getNamesOfComing(sheet, COMING_RANGE) -> list:
    coming = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=COMING_RANGE).execute()
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
    names = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=Constants.MASTER_NAME_RANGE).execute()
    names_values = names.get('values', [])

    if not names_values:
        print('Tracking Sheet Name Empty!')
        quit()

    filtered_list = []
    for (idx, names_value) in enumerate(names_values):
        if idx in coming_list:
            filtered_list.append(names_value)

    return filtered_list


def getCurrAttendance(sheet) -> list:
    currAttendance = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=Constants.ATTENDANCE_SHEET_NAME_RANGE).execute()
    currAttendance_values = currAttendance.get('values', [])

    if not currAttendance_values:
        print('Attendance Sheet Name!')
        quit()

    return currAttendance_values


def checkAttendanceSheet(sheet, currAttendance, day, coming_list) -> None:
    trainingDates = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=Constants.ATTENDANCE_SHEET_DATE_RANGE).execute()
    trainingDates_values = trainingDates.get('values', [])[0]

    if not trainingDates_values:
        print('Attendance Sheet Date!')
        quit()

    dateIdx = chr(trainingDates_values.index(day) + 5 + 65)
        
    currAttendance_telegram_values = []
    currAttendance_telegram_values.append([i[Constants.ATTENDANCE_TELE_ID].lower() for i in currAttendance])
    currAttendance_telegram_values = currAttendance_telegram_values[0]

    coming_telegram_values = []
    coming_telegram_values.append([i[Constants.MASTER_ID_TELE_LIST].lower() for i in coming_list])
    coming_telegram_values = coming_telegram_values[0]

    currAttendanceSize = len(currAttendance_telegram_values)

    batch_update = []

    for coming_telegram_value in coming_telegram_values:
        # Add their name and telegram handle if not in attendance sheet
        if coming_telegram_value not in currAttendance_telegram_values:
            currAttendanceSize += 1
            currAttendanceName = ""
            currAttendanceYear = ""
            currAttendanceTele = ""

            for coming_list_value in coming_list:
                if coming_telegram_value == coming_list_value[Constants.MASTER_ID_TELE_LIST].lower():
                    currAttendanceName = coming_list_value[Constants.MASTER_ID_NAME_LIST]
                    currAttendanceYear = coming_list_value[Constants.MASTER_ID_YEAR_LIST]
                    currAttendanceTele = coming_telegram_value

                    break
            
            new_attendance_range = Constants.ATTENDANCE_SHEET_NAME + '!A' + str(currAttendanceSize + Constants.ATTENDANCE_DATA_START)
            new_attendance_row = [currAttendanceSize, currAttendanceName, currAttendanceTele, Constants.ATTENDANCE_MEMBER_ROLE, currAttendanceYear]

            currAttendance.append(new_attendance_row)
            batch_update.append({"range": new_attendance_range, "values": [new_attendance_row]})

            print("Added " + currAttendanceTele + " at row " + str(new_attendance_row))
        
            coming_attendance_row = str(currAttendanceSize + Constants.ATTENDANCE_DATA_START)
            coming_attendance_range = Constants.ATTENDANCE_SHEET_NAME + '!' + dateIdx + str(coming_attendance_row)
            coming_attendance_value = '0'

            batch_update.append({"range": coming_attendance_range, "values": [[coming_attendance_value]]})
            print("Updated " + coming_telegram_value + " at row " + str(coming_attendance_row))
        
        else:
            coming_attendance_row = currAttendance_telegram_values.index(coming_telegram_value) + Constants.ATTENDANCE_DATA_START + 1
            coming_attendance_range = Constants.ATTENDANCE_SHEET_NAME + '!' + dateIdx + str(coming_attendance_row)
            coming_attendance_value = '0'

            batch_update.append({"range": coming_attendance_range, "values": [[coming_attendance_value]]})
            print("Updated " + coming_telegram_value + " at row " + str(coming_attendance_row))

    update_values_request = {
        "valueInputOption": "USER_ENTERED",
        "data": batch_update,
    }

    sheet.values().batchUpdate(
        spreadsheetId=Constants.SPREADSHEET_ID_MASTER,
        body=update_values_request,
    ).execute()

    return currAttendance

def main() -> None:
    creds = auth()

    nextMon = next_weekday(0)

    try:
        service = build('sheets', 'v4', credentials=creds)
        # Call the Sheets API
        sheet = service.spreadsheets()

        (monIdx, wedIdx) = parseColumnByDates(sheet, nextMon)

        MEMBER_TRACKING_SHEET_COMING_MONDAY_RANGE = Constants.MASTER_SHEET_NAME + '!' + monIdx + Constants.MASTER_DATA_START + ':' + monIdx
        MEMBER_TRACKING_SHEET_COMING_WEDNESDAY_RANGE = Constants.MASTER_SHEET_NAME + '!' + wedIdx + Constants.MASTER_DATA_START + ':' + wedIdx

        monNames = getNamesOfComing(sheet, MEMBER_TRACKING_SHEET_COMING_MONDAY_RANGE)
        wedNames = getNamesOfComing(sheet, MEMBER_TRACKING_SHEET_COMING_WEDNESDAY_RANGE)

        monNames_filtered = filterComing(sheet, monNames)
        wedNames_filtered = filterComing(sheet, wedNames)

        currAttendance = getCurrAttendance(sheet)

        currAttendance = checkAttendanceSheet(sheet, currAttendance, str(nextMon.day), monNames_filtered)
        currAttendance = checkAttendanceSheet(sheet, currAttendance, str(nextMon.day + 2), wedNames_filtered)

        print("Finished! Cleaning up...")

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()