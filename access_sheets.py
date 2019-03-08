#
# access_sheets.py -- access Google spreadsheets using the Google API's
#
# @author Google
# @modified Mark Sattolo <epistemik@gmail.com>
# @revised 2019-03-08
#

from __future__ import print_function

import pickle
import os.path as osp
import json
import datetime as dt

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from mhs_google_config import *

SPREADSHEET_ID = BUDGET_QTRLY_SHEET
SHEET_RANGE    = BUDGET_QTRLY_SAMPLE_RANGE
CURRENT_SCOPE  = SHEETS_RW_SCOPE
TOKEN = SHEETS_EPISTEMIK_RW_TOKEN

now = dt.datetime.strftime(dt.datetime.now(), "%Y-%m-%d_%H-%M-%S")


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
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', CURRENT_SCOPE)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open(TOKEN, 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    srv_sheets = service.spreadsheets()

    def retrieve_sheet():
        # Do a document "get" request and print the results as formatted JSON
        sheet = srv_sheets.get(spreadsheetId = SPREADSHEET_ID).execute()
        # print(json.dumps(document, indent=4))

        sheet_title = sheet.get('properties').get('title')
        if sheet_title is None:
            sheet_title = 'MHS_Spreadsheet'
        print("The title of the spreadsheet is: {}".format(sheet_title))

        # print spreadsheet as json file -- add a timestamp to get a unique file name
        out_file = sheet_title + '.' + now + ".json"
        print("out_file is '{}'".format(out_file))
        fp = open(out_file, 'w')
        json.dump(sheet, fp, indent=4)

    def retrieve_range():
        result = srv_sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=SHEET_RANGE).execute()
        print("after result")
        values = result.get('values', [])
        print("after values")

        if not values:
            print('No data found.')
        else:
            print("after else")
            # json.dumps(values, indent=4)
            # print('Name, Major:')
            for row in values:
                # Print columns A and E, which correspond to indices 0 and 4.
                # print('%s, %s' % (row[0], row[4]))
                print(str(row))

    # retrieve_sheet()
    retrieve_range()

    print('PROGRAM ENDED.')


if __name__ == '__main__':
    main()
