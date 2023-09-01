# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a master tracking spreadsheet.
SPREADSHEET_ID_MASTER = '1Wvk4UGKgiT_UACsgzUSepydVA7GgO5FbqtsL9p3qM-4'
SPREADSHEET_ID_FORM = '1eEHQB_kuZ0mFfq97kQkPhCSWpli723Szp7TrJWBP3xU'

MASTER_SHEET_NAME = 'Sem 1'
MASTER_DATA_START = '19'

MASTER_ID_NAME = 'A'
MASTER_ID_YEAR = 'B'
MASTER_ID_TELE = 'C'

# [[Date]]
MASTER_DATE_RANGE = MASTER_SHEET_NAME + '!E2:AF2'
# [[Full Name, Telegram Display name, Telegram Handle]]
MASTER_NAME_RANGE = MASTER_SHEET_NAME + '!A' + MASTER_DATA_START + ':' + MASTER_ID_TELE

FORM_SHEET_NAME = 'Week 4'
FORM_START = '2'

FORM_NAME = 'B'
FORM_DATE = 'E'

FORM_NAME_ID = 0
FORM_YEAR_ID = 1
FORM_TELE_ID = 2
FORM_DATE_ID = 3

FORM_DATA_RANGE = FORM_SHEET_NAME + '!' + FORM_NAME + FORM_START + ':' + FORM_DATE

MON = 'Monday'
WED = 'Wednesday'
BOTH = 'Both Days'

PAID = 'PAID'

TOKEN = 'token.json'
CREDENTIALS = 'credentials.json'