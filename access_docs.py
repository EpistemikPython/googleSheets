#
# access_docs.py -- access Google documents using the Google API's
#
# @author Google
# @modified Mark Sattolo <epistemik@gmail.com>
# @revised 2019-03-02
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

DOCUMENT_ID = READING_DOC
CURRENT_SCOPE = DOCS_RW_SCOPE
TOKEN = DOCS_EPISTEMIK_RW_TOKEN

now = dt.datetime.strftime(dt.datetime.now(), "%Y-%m-%d_%H-%M-%S")


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
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, CURRENT_SCOPE)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open(TOKEN, 'wb') as token:
            pickle.dump(creds, token)

    service = build('docs', 'v1', credentials=creds)

    # Do a document "get" request and print the results as formatted JSON
    document = service.documents().get(documentId = DOCUMENT_ID).execute()
    # print(json.dumps(document, indent=4))

    doc_title = document.get('title')
    print("The title of the document is: {}".format(doc_title))
    # print('The body of the document is: {}'.format(document.get('body')))

    # print document as json file -- add a timestamp to get a unique file name
    out_file = doc_title + '.' + now + ".json"
    print("out_file is '{}'".format(out_file))
    fp = open(out_file, 'w')
    json.dump(document, fp, indent=4)


if __name__ == '__main__':
    main()
