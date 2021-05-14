#
# modify_docs.py -- modify Google documents using the Google API's
#
# @author Google
# @modified Mark Sattolo <epistemik@gmail.com>
# @revised 2019-03-12
# @version Python3.6
#

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
TOKEN = DOCS_EPISTEMIK_RW_TOKEN['P4']

now = dt.datetime.strftime(dt.datetime.now(), "%Y-%m-%dT%H-%M-%S")

# simple test data
extra_text = [
    {
        'insertText': {
            'location': {
                'index': 27,
            },
            'text': "\tForeign Secretary for the Commonwealth\n\n\n\n"
        }
    },
    {
        'insertText': {
            'location': {
                'index': 27,
            },
            'text': "\tauthor of Paradise Lost\n"
        }
    },
    {
        'insertText': {
            'location': {
                'index': 27,
            },
            'text': "\t1608 - 1674\n"
        }
    },
    {
        'insertText': {
            'location': {
                'index': 27,
            },
            'text': "\tJohn Milton:\n"
        }
    },
    {
        'insertText': {
            'location': {
                'index': 27,
            },
            'text': "\t{}\n".format(now)
        }
    }
]

# example text format changes
text_style_updates = [
    {
        'updateTextStyle': {
            'range': {
                'startIndex': 27,
                'endIndex': 48
            },
            'textStyle': {
                'bold': True,
                'italic': True
            },
            'fields': 'bold,italic'
        }
    },
    {
        'updateTextStyle': {
            'range': {
                'startIndex': 48,
                'endIndex': 62
            },
            'textStyle': {
                'weightedFontFamily': {
                    'fontFamily': 'Times New Roman'
                },
                'fontSize': {
                    'magnitude': 14,
                    'unit': 'PT'
                },
                'foregroundColor': {
                    'color': {
                        'rgbColor': {
                            'blue': 0.0,
                            'green': 1.0,
                            'red': 0.0
                        }
                    }
                }
            },
            'fields': 'foregroundColor,weightedFontFamily,fontSize'
        }
    },
    {
        'updateTextStyle': {
            'range': {
                'startIndex': 62,
                'endIndex': 75
            },
            'textStyle': {
                'link': {
                    'url': 'www.example.com'
                }
            },
            'fields': 'link'
        }
    }
]

# example paragraph style changes
pgraph_style_updates = [
    {
        'updateParagraphStyle': {
            'range': {
                'startIndex': 75,
                'endIndex': 100
            },
            'paragraphStyle': {
                'namedStyleType': 'HEADING_1',
                'spaceAbove': {
                    'magnitude': 10.0,
                    'unit': 'PT'
                },
                'spaceBelow': {
                    'magnitude': 10.0,
                    'unit': 'PT'
                }
            },
            'fields': 'namedStyleType,spaceAbove,spaceBelow'
        }
    },
    {
        'updateParagraphStyle': {
            'range': {
                'startIndex': 100,
                'endIndex': 140
            },
            'paragraphStyle': {
                'borderLeft': {
                    'color': {
                        'color': {
                            'rgbColor': {
                                'blue': 0.0,
                                'green': 0.0,
                                'red': 1.0
                            }
                        }
                    },
                    'dashStyle': 'DASH',
                    'padding': {
                        'magnitude': 20.0,
                        'unit': 'PT'
                    },
                    'width': {
                        'magnitude': 15.0,
                        'unit': 'PT'
                    },
                }
            },
            'fields': 'borderLeft,borderRight'
        }
    }
]


def modify_docs_main():
    """
    Modify the text and/or styles of a specified Google document
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

    service = build('docs', 'v1', credentials=creds)

    # send the extra text
    t_reply = service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': extra_text}).execute()
    # show the reply message
    print("The reply of the document update operation is: {}".format(json.dumps(t_reply, indent=4)))

    # send the text formatting changes
    f_reply = service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': text_style_updates}).execute()
    # show the reply message
    print("The reply of the text formatting operation is: {}".format(json.dumps(f_reply, indent=4)))

    # send the paragraph formatting changes
    p_reply = service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': pgraph_style_updates}).execute()
    # show the reply message
    print("The reply of the paragraph formatting operation is: {}".format(json.dumps(p_reply, indent=4)))

    print('\n >>> PROGRAM ENDED.')


if __name__ == '__main__':
    modify_docs_main()
