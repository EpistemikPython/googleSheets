#
# access_docs.py -- access Google documents using the Google API's
#
# @author Google
# @modified Mark Sattolo <epistemik@gmail.com>
# @revised 2019-03-02
#
"""
from __future__ import print_function
from apiclient import discovery
from httplib2 import Http
from oauth2client import client
from oauth2client import file
from oauth2client import tools
import json

# Set doc ID, as found at `https://docs.google.com/document/d/YOUR_DOC_ID/edit`
DOCUMENT_ID=YOUR_DOC_ID

# Set the scopes and discovery info
SCOPES = 'https://www.googleapis.com/auth/documents.readonly'
DISCOVERY_DOC = 'https://docs.googleapis.com/$discovery/rest?version=v1&key=<YOUR_API_KEY>'

# Initialize credentials and instantiate Docs API service
store = file.Storage('token.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
    creds = tools.run_flow(flow, store)
service = discovery.build('docs', 'v1', http=creds.authorize(
    Http()), discoveryServiceUrl=DISCOVERY_DOC)

# Do a document "get" request and print the results as formatted JSON
result = service.documents().get(documentId=DOCUMENT_ID).execute()
print(json.dumps(result, indent=4, sort_keys=True))
"""

import pickle
import os.path as osp
import json
import datetime as dt

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from mhs_google_config import *

DOCUMENT_ID = READING_DOC
CURRENT_SCOPE = DOCS_RW_SCOPE
TOKEN = DOCS_EPISTEMIK_RW_TOKEN

now = dt.datetime.strftime(dt.datetime.now(), "%Y-%m-%d_%H-%M-%S")

# simple test data
requests = [
    {
        'insertText': {
            'location': {
                'index': 28,
            },
            'text': "\tauthor of Paradise Lost\n"
        }
    },
    {
        'insertText': {
            'location': {
                'index': 28,
            },
            'text': "\t1608 - 1674\n"
        }
    },
    {
        'insertText': {
            'location': {
                'index': 28,
            },
            'text': "\tJohn Milton:\n"
        }
    },
    {
        'insertText': {
            'location': {
                'index': 28,
            },
            'text': "\t{}\n".format(now)
        }
    }
]


def main():
    """
    Get the title and/or body of a specified Google document
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
            flow = InstalledAppFlow.from_client_secrets_file('secrets/credentials.json', CURRENT_SCOPE)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open(TOKEN, 'wb') as token:
            pickle.dump(creds, token)

    service = build('docs', 'v1', credentials=creds)

    # send the modifications
    document = service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': requests}).execute()
    print("The return of the document update operation is: {}".format(json.dumps(document, indent=4)))


if __name__ == '__main__':
    main()
