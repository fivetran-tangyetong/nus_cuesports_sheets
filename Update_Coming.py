from __future__ import print_function

import os.path

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import Constants
import utils

spreadsheet_dates = []
spreadsheet_names = []

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
    creds = utils.auth()

    nextMon = utils.next_weekday(0)
    nextWed = utils.next_weekday(2)

    coming_mon = getComing(Constants.TELE_MON)
    coming_wed = getComing(Constants.TELE_WED)

    try:
        service = build('sheets', 'v4', credentials=creds)
        # Call the Sheets API
        sheet = service.spreadsheets()

        monIdx = utils.parseColumnByDate(sheet, nextMon)
        wedIdx = utils.parseColumnByDate(sheet, nextWed)
        
        if monIdx != '':
            updateMembers(sheet, coming_mon, monIdx)
            MEMBER_TRACKING_SHEET_COMING_MONDAY_RANGE = Constants.MASTER_SHEET_NAME + '!' + monIdx + Constants.MASTER_DATA_START + ':' + monIdx
            checkAllMembers(sheet, coming_mon, MEMBER_TRACKING_SHEET_COMING_MONDAY_RANGE)

        if wedIdx != '':
            updateMembers(sheet, coming_wed, wedIdx)
            MEMBER_TRACKING_SHEET_COMING_WEDNESDAY_RANGE = Constants.MASTER_SHEET_NAME + '!' + wedIdx + Constants.MASTER_DATA_START + ':' + wedIdx
            checkAllMembers(sheet, coming_wed, MEMBER_TRACKING_SHEET_COMING_WEDNESDAY_RANGE)
                

    except HttpError as err:
        print(err)
        
    print("Finished! Cleaning up...")


if __name__ == '__main__':
    main()