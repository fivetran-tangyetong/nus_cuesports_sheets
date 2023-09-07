from __future__ import print_function

import os.path
import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import Constants

DISPLAY = """
 ├ sihui
├ Team
├ John Gao Jiahao
├ Yi Long A
├ Jonathan Wee
├ Joshua Tan
├ bryan
├ Gan Kai Xuan
├ Justin
├ Opal
├ Hui Buen
├ lavonne
├ Brendan Ng
├ gene
├ An Jo
├ felixx
├ Eliot
├ KahJyun
├ Lim Jun Ming
├ Yi Lin
├ Issac
├ Rajarshi Basu
├ arthur
├ Arvind
├ Roderick
├ charl
├ Yu Heng
├ Anissa
├ Weiqing W
├ Xiaowei Shu
├ Ang Zhi Tat
├ hbs
├ Kaushik Raman
├ Marta
├ Lim Zhong Zhi
├ Vimuth
├ tzhenyao
├ Eli
├ David
├ Simin
├ Sze Perng
├ Feng Yuan
├ pahaul
└ chun wei lee
"""

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


def getAllMembers(sheet) -> list:
    names = sheet.values().get(spreadsheetId=Constants.SPREADSHEET_ID_MASTER, range=Constants.MASTER_NAME_RANGE).execute()
    names_values = names.get('values', [])

    return names_values

def main() -> None:
    creds = auth()

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