from __future__ import print_function

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import Constants
import utils

DISPLAY = """
├ Eliot
├ tzhenyao
├ Yap Bing Feng
├ Andy
├ Zheng Jie
├ David
├ John Gao Jiahao
├ bryan
├ Joven Chua
├ Joshua Tan
├ Eli
├ jessie
├ Jun Hern
├ Jonathan Wee
├ Duc
├ Marcus Wong
├ keefe
├ Neo
├ Ang Zhi Tat
├ gene
├ linhns
├ An Jo
├ Dominic
├ Marta
├ Sze Perng
├ Yi Long A
└ Yu Heng

"""

def getAllMembers(sheet) -> list:
    names = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=Constants.MASTER_NAME_RANGE).execute()
    names_values = names.get('values', [])

    return names_values

def main() -> None:
    creds = utils.auth()

    try:
        service = build('sheets', 'v4', credentials=creds)
        # Call the Sheets API
        sheet = service.spreadsheets()

        names_values = getAllMembers(sheet)

        telegram_values = [i[Constants.MASTER_ID_DISP_LIST] for i in names_values]

        display_filtered = DISPLAY.strip().split("\n")

        final = []

        for display_filtered_value in display_filtered:

            if display_filtered_value.startswith('├ ') or display_filtered_value.startswith('└ '):
                display_name = display_filtered_value[2:].strip()
            else:
                display_name = display_filtered_value.strip()

            display_name_num = telegram_values.count(display_name)

            if display_name_num > 1:
                print("Multiple Display Name for: " + display_name)
            elif display_name_num == 0:
                print("Display Name DNE for: " + display_name)
            else:
                final.append(names_values[telegram_values.index(display_name)][Constants.MASTER_ID_TELE_LIST])

            if len(final) == 36:
                break
        
        print("Printing " + str(len(final)) + " Entries!")

        for i in final:
            print(i)

    except HttpError as err:
        print(err)
        
    print("Finished! Cleaning up...")


if __name__ == '__main__':
    main()