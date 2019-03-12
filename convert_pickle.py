#
# convert_pickle.py -- convert credentials tokens from pickle protocol 3 to protocol 4
#
# @author Google
# @modified Mark Sattolo <epistemik@gmail.com>
# @revised 2019-03-12
# @version Python3.6
#

import pickle
import os.path as osp
from googleapiclient.discovery import build

TOKEN = 'token.pickle'
HIGHEST_TOKEN = 'token.pickle4'


def main():
    creds = None
    # open the existing token
    if osp.exists(TOKEN):
        with open(TOKEN, 'rb') as token:
            # ?? will not load protocol 2
            # see out/convert_pickle_v2-to-v4.error.term
            creds = pickle.load(token)

    # save at the highest protocol
    with open(HIGHEST_TOKEN, 'wb') as h_token:
        pickle.dump(creds, h_token, pickle.HIGHEST_PROTOCOL)

    # service = build('docs', 'v1', credentials=creds)
    service = build('sheets', 'v4', credentials=creds)
    print("service is type '{}'".format(type(service)))

    print('PROGRAM ENDED.')


if __name__ == '__main__':
    main()
