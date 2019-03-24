#
# access_sheets.py -- access Google spreadsheets using the Google API's
#
# @author Google
# @modified Mark Sattolo <epistemik@gmail.com>
# @revised 2019-03-23
# @version Python3.6


import pickle
import os.path as osp
import datetime as dt
import json

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from mhs_google_config import *

SPREADSHEET_ID = BUDGET_QTRLY_SPD_SHEET
SRC_SHEET_ID   = BUDQTR_ALL_INC_SHEET
DEST_SHEET_ID  = BUDQTR_CALCULNS_SHEET

COPY_SHEET_RANGE  = BUDGET_QTRLY_READ_RANGE
WRITE_SHEET_RANGE = BUDGET_QTRLY_WRITE_RANGE
WRITE_CELL_RANGE = 'Calculations!P36'

CURRENT_SCOPE  = SHEETS_RW_SCOPE
TOKEN = SHEETS_EPISTEMIK_RW_TOKEN['P4']

now = dt.datetime.strftime(dt.datetime.now(), "%Y-%m-%dT%H-%M-%S")

# write multiple values to a row
write_range_body = {
    'values': [
        [
            '21', 'DE', '=3+9', '32.654', '$56.87'
        ]
        # Additional rows ...
    ]
}

# update the value of a single cell
write_cell_body = {
    'values': [
        [
            '$999.99'
        ]
    ]
}

# copy and paste a range of values
copy_paste_request = {
    "copyPaste": {
        "source": {
            "sheetId": SRC_SHEET_ID,
            "startRowIndex": 67,
            "endRowIndex"  : 76,
            "startColumnIndex":  0,
            "endColumnIndex"  : 17
        },
        "destination": {
            "sheetId": DEST_SHEET_ID,
            "startRowIndex": 43,
            "endRowIndex"  : 52,
            "startColumnIndex":  0,
            "endColumnIndex"  : 17
        },
        "pasteType": "PASTE_NORMAL",
        "pasteOrientation": "NORMAL"
    }
}

# copy and paste body
copy_body = {
    'requests': copy_paste_request
}

# cut and paste a range
cut_paste_request = [
    {
        "cutPaste": {
            "source": {
                "sheetId": SRC_SHEET_ID,
                "startRowIndex": 67,
                "endRowIndex"  : 76,
                "startColumnIndex":  0,
                "endColumnIndex"  : 17
            },
            "destination": {
                "sheetId": DEST_SHEET_ID,
                "rowIndex"    : 32,
                "columnIndex" :  0
            },
            "pasteType": "PASTE_NORMAL"
        }
    }
    # can add other types of request to this list
]

# cut and paste body
cut_body = {
    'requests': cut_paste_request
}

requests = []
# Change the spreadsheet's title.
requests.append({
    'updateSpreadsheetProperties': {
        'properties': {
            'title': 'NEW TITLE'
        },
        'fields': 'title'
    }
})
# Find and replace text
requests.append({
    'findReplace': {
        'find': 'X',
        'replacement': 'Z',
        'allSheets': True
    }
})
# Add additional requests (operations) ...

body = {
    'requests': requests
}

# write values to two separate cells
data = [
    {
        'range': 'Calculations!P47',
        'values': [
            [
                '$7022.77'
            ],
            # Additional rows
        ]
    },
    {
        'range': 'Calculations!P49',
        'values': [
            [
                '$8044.66'
            ],
            # Additional rows
        ]
    },
    # Additional ranges to update ...
]


def modify_sheets_main():
    """
    Copy a range of cells from one sheet to another, write new values &
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if osp.exists(TOKEN):
        with open(TOKEN, 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, CURRENT_SCOPE)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open(TOKEN, 'wb') as token:
            pickle.dump(creds, token, pickle.HIGHEST_PROTOCOL)

    # Call the Sheets API
    service = build('sheets', 'v4', credentials=creds)
    srv_sheets = service.spreadsheets()

    def cut_and_paste():
        response = srv_sheets.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=cut_body).execute()

        for item in response.get('replies'):
            print("{}\n\n".format(item))
        print(response)

    def copy_and_paste():
        response = srv_sheets.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=copy_body).execute()

        for item in response.get('replies'):
            print("{}\n\n".format(item))
        print(response)

    def find_and_replace():
        response = srv_sheets.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()

        find_replace_response = response.get('replies')[1].get('findReplace')
        print('{0} replacements made.'.format(find_replace_response.get('occurrencesChanged')))

    def write_to_range():
        vals = srv_sheets.values()
        result = vals.update(spreadsheetId=SPREADSHEET_ID, range=WRITE_SHEET_RANGE,
                             valueInputOption='USER_ENTERED', body=write_range_body).execute()
        print("{} cells updated.".format(result.get('updatedCells')))

    def write_to_cell():
        vals = srv_sheets.values()
        result = vals.update(spreadsheetId=SPREADSHEET_ID, range=WRITE_CELL_RANGE,
                             valueInputOption='USER_ENTERED', body=write_cell_body).execute()
        print("{} cells updated.".format(result.get('updatedCells')))

    def write_two_cells():
        my_body = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
        }
        vals = srv_sheets.values()
        result = vals.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=my_body).execute()

        print('{} cells updated!'.format(result.get('totalUpdatedCells')))
        print(json.dumps(result, indent=4))

    # cut_and_paste()
    # copy_and_paste()
    # find_and_replace()
    # write_to_range()
    # write_to_cell()
    write_two_cells()

    print('\nPROGRAM ENDED.')


if __name__ == '__main__':
    modify_sheets_main()
