#
# createGnucashTxs.py -- parse a Monarch record, possibly from a json file,
#                        create Gnucash transactions from the data
#                        and write to a Gnucash file
#
# Copyright (c) 2018, 2019 Mark Sattolo <epistemik@gmail.com>
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License as
# published by the Free Software Foundation; either version 2 of
# the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# @author Mark Sattolo <epistemik@gmail.com>
#

from __future__ import print_function

import pickle
import os.path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from mhs_google_config import *

DOCUMENT_ID = BUDGET_QTRLY_SHEET
CURRENT_SCOPE = SHEETS_RW_SCOPE
CLIENT_SECRETS = 'secrets/credentials.json'

def main():
    """
    Shows basic usage of the Docs API.
    Prints the title of a sample document.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS, CURRENT_SCOPE)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # service = build('docs', 'v1', credentials=creds)
    # Retrieve the documents contents from the Docs service.
    # document = service.documents().get(documentId = DOCUMENT_ID).execute()

    service = build('sheets', 'v4', credentials=creds)
    # Retrieve the documents contents from the Docs service.
    sheet = service.spreadsheets().get(spreadsheetId = DOCUMENT_ID).execute()

    print('The title of the document is: {}'.format(document.get('title')))
    # print('The body of the document is: {}'.format(document.get('body')))


if __name__ == '__main__':
    main()
