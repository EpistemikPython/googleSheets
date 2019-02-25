#
# access_sheets.py -- access Google spreadsheets using the Google API's
#
# @author Google
# @revised Mark Sattolo <epistemik@gmail.com>
#

from __future__ import print_function

import pickle
import os.path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from mhs_google_config import *

SPREADSHEET_ID = SAMPLE_SPREADSHEET
SHEET_RANGE    = SAMPLE_RANGE

CURRENT_SCOPE = SHEETS_RW_SCOPE
TOKEN = SHEETS_SAMPLE_RW_TOKEN


def main():
    """
    Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists(TOKEN):
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
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=SHEET_RANGE).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            print('%s, %s' % (row[0], row[4]))


if __name__ == '__main__':
    main()
