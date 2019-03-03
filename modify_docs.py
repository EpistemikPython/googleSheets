#
# modify_docs.py -- update Google documents using the Google API's
#
# @author Google
# @modified Mark Sattolo <epistemik@gmail.com>
# @revised 2019-03-02
#

import pickle
import os.path
import json

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from mhs_google_config import *

DOCUMENT_ID = READING_DOC
CURRENT_SCOPE = DOCS_RW_SCOPE
TOKEN = DOCS_EPISTEMIK_RW_TOKEN


def authorize(creds):
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


def main():
    """
    Update a specified Google document
    """
    creds = None
    authorize(creds)

    # simple test data
    requests = [
         {
            'insertText': {
                'location': {
                    'index': 28,
                },
                'text': "author of Paradise Lost\n"
            }
        },
                 {
            'insertText': {
                'location': {
                    'index': 28,
                },
                'text': "1608 - 1674\n"
            }
        },
                 {
            'insertText': {
                'location': {
                    'index': 28,
                },
                'text': "John Milton:\n"
            }
        },
    ]

    service = build('docs', 'v1', credentials=creds)
    # send the modifications
    result = service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': requests}).execute()

    print("The return of the document update operation is: {}".format(result))


if __name__ == '__main__':
    main()
