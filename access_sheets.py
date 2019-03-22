#
# access_sheets.py -- access Google spreadsheets using the Google API's
#
# @author Google
# @modified Mark Sattolo <epistemik@gmail.com>
# @revised 2019-03-12
# @version Python3.6


import pickle
import os.path as osp
import json
import datetime as dt

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from mhs_google_config import *

SPREADSHEET_ID = BUDGET_QTRLY_SPD_SHEET
SHEET_RANGE    = BUDGET_QTRLY_SAMPLE_RANGE
CURRENT_SCOPE  = SHEETS_RW_SCOPE
TOKEN = SHEETS_EPISTEMIK_RW_TOKEN['P4']

now = dt.datetime.strftime(dt.datetime.now(), "%Y-%m-%dT%H-%M-%S")


def main():
    """
    Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
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

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    srv_sheets = service.spreadsheets()
    sheet = srv_sheets.get(spreadsheetId=SPREADSHEET_ID).execute()
    props = sheet.get('properties')

    sheet_title = props.get('title')
    if sheet_title is None:
        sheet_title = 'MHS_Spreadsheet'

    def save_sheet():
        # print spreadsheet as json file -- add a timestamp to get a unique file name
        out_file = sheet_title + '.' + now + ".json"
        print("\nSheet out_file is '{}'".format(out_file))
        sfp = open(out_file, 'w')
        json.dump(sheet, sfp, indent=4)
        sfp.close()

    def save_props():
        # this DOES NOT WORK!?
        # props = srv_sheets.properties().get(spreadsheetId = SPREADSHEET_ID).execute()

        props_title = sheet_title + '-properties'

        # print properties as json file -- add a timestamp to get a unique file name
        out_file = props_title + '.' + now + ".json"
        print("\nProps out_file is '{}'".format(out_file))
        pfp = open(out_file, 'w')
        json.dump(props, pfp, indent=4)
        pfp.close()

    def retrieve_range():
        result = srv_sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=SHEET_RANGE).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
        else:
            print("\nRange [{}]:".format(SHEET_RANGE))
            for row in values:
                print(str(row))

    save_sheet()
    save_props()
    retrieve_range()

    print('\nPROGRAM ENDED.')


if __name__ == '__main__':
    main()
