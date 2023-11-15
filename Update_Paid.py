from __future__ import print_function

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import Constants
import utils


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


def updateMembersPaid(sheet, paid_list, mon_idx, wed_idx) -> None:
    print("Updating Members Paid!")
    names = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=Constants.MASTER_NAME_RANGE).execute()
    names_values = names.get('values', [])

    if not names_values:
        print('Tracking Sheet Name Empty!')
        quit()

    telegram_values = [i[Constants.MASTER_ID_TELE_LIST].lower().strip() for i in names_values]
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
    creds = utils.auth()

    nextMon = utils.next_weekday(0)
    nextWed = utils.next_weekday(2)

    try:
        service = build('sheets', 'v4', credentials=creds)
        # Call the Sheets API
        sheet = service.spreadsheets()

        paid_list = getMembersPaid(sheet)
        monIdx = utils.parseColumnByDate(sheet, nextMon)
        wedIdx = utils.parseColumnByDate(sheet, nextWed)
        
        updateMembersPaid(sheet, paid_list, monIdx, wedIdx)

    except HttpError as err:
        print(err)
        
    print("Finished! Cleaning up...")


if __name__ == '__main__':
    main()