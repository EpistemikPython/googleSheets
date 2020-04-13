##############################################################################################################################
# coding=utf-8
#
# google_utilities.py -- useful constants, functions & classes for Google
#
# some code from Google quickstart examples
#
# Copyright (c) 2020 Mark Sattolo <epistemik@gmail.com>

__author__         = 'Mark Sattolo'
__author_email__   = 'epistemik@gmail.com'
__google_api_python_client_version__ = '1.7.11'
__created__ = '2019-04-07'
__updated__ = '2020-04-13'

import threading
from sys import path
import os.path as os_path
import pickle5 as pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from typing import Union
path.append("/home/marksa/dev/git/Python/Utilities/")
from investment import *

# see https://github.com/googleapis/google-api-python-client/issues/299
# build('drive', 'v3', http=http, cache_discovery=False)
lg.getLogger('googleapiclient.discovery_cache').setLevel(lg.ERROR)

# sheet names in Budget Quarterly
ML_WORK_SHEET:str  = 'ML Work'
CALCULNS_SHEET:str = 'Calculations'

# first data row in the sheets
BASE_ROW:int = 3

CREDENTIALS_FILE:str = 'secrets/credentials.json'

SHEETS_RW_SCOPE:list = ['https://www.googleapis.com/auth/spreadsheets']

SHEETS_EPISTEMIK_RW_TOKEN:dict = {
    'P2' : 'secrets/token.sheets.epistemik.rw.pickle2' ,
    'P3' : 'secrets/token.sheets.epistemik.rw.pickle3' ,
    'P4' : 'secrets/token.sheets.epistemik.rw.pickle4' ,
    'P5' : 'secrets/token.sheets.epistemik.rw.pickle5'
}
GGL_SHEETS_TOKEN:str = SHEETS_EPISTEMIK_RW_TOKEN['P5']

# Spreadsheet ID
BUDGET_QTRLY_ID_FILE:str = 'secrets/Budget-qtrly.id'

# sheet names in Budget Quarterly
ALL_INC_SHEET:str    = 'All Inc 1'
ALL_INC_2_SHEET:str  = 'All Inc 2'
NEC_INC_SHEET:str    = 'Nec Inc 1'
NEC_INC_2_SHEET:str  = 'Nec Inc 2'
QTR_ASTS_SHEET:str   = 'Assets 1'
QTR_ASTS_2_SHEET:str = 'Assets 2'
BAL_1_SHEET:str      = 'Balance 1'
BAL_2_SHEET:str      = 'Balance 2'

FILL_CELL_VAL = Union[str, Decimal]


# noinspection PyUnresolvedReferences
def get_credentials(logger:lg.Logger=None) -> pickle:
    """get the proper credentials needed to write to the Google spreadsheet"""
    if logger: logger.info(get_current_time())
    creds = None
    if os_path.exists(GGL_SHEETS_TOKEN):
        with open(GGL_SHEETS_TOKEN, 'rb') as token:
            creds = pickle.load(token)

    # if there are no (valid) credentials available, let the user log in.
    if creds is None or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SHEETS_RW_SCOPE)
            creds = flow.run_local_server()
        # save the credentials for the next run
        with open(GGL_SHEETS_TOKEN, 'wb') as token:
            pickle.dump(creds, token, pickle.DEFAULT_PROTOCOL)

    return creds


class GoogleUpdate:
    # prevent different instances/threads from writing at the same time
    _lock = threading.Lock()

    def __init__(self, p_logger:lg.Logger):
        self._lgr = p_logger
        self._data = list()
        self._lgr.info(F"\n\tLaunch {self.__class__.__name__} instance with lock {str(self._lock)}\n\t"
                       F" at Runtime = {get_current_time()}\n")

    def get_data(self) -> list:
        return self._data

    # noinspection PyAttributeOutsideInit
    def begin_session(self):
        # CANNOT have a separate Session on the Google file
        self._lock.acquire()
        self._lgr.info(F"acquired lock at {get_current_time()}")

        creds = get_credentials()
        service = build('sheets', 'v4', credentials = creds, cache_discovery = False)
        self.vals = service.spreadsheets().values()

    def end_session(self):
        # RELEASE the Session on the Google file
        self._lock.release()
        self._lgr.debug(F"released lock at {get_current_time()}")

    def __get_budget_id(self) -> str:
        """
        get the budget id string from the file in the secrets folder
        """
        with open(BUDGET_QTRLY_ID_FILE, "r") as bfp:
            fid = bfp.readline().strip()
        self._lgr.debug(get_current_time() + F" / __get_budget_id(): Budget Id = {fid}\n")
        return fid

    def fill_cell(self, sheet:str, col:str, row:int, val:FILL_CELL_VAL):
        """
        create the information to update a Google Sheets cell and add to the data list
        :param sheet: particular sheet in my Google spreadsheet to update
        :param   col: column to update
        :param   row: to update
        :param   val: str OR Decimal: value to fill with
        """
        self._lgr.debug(get_current_time())

        value = val.to_eng_string() if isinstance(val, Decimal) else val
        cell = {'range': sheet + '!' + col + str(row), 'values': [[value]]}
        self._lgr.debug(F"fill_cell() = {cell}\n")
        self._data.append(cell)

    # noinspection PyTypeChecker
    def send_sheets_data(self) -> dict:
        """
        Send the data list to my Google sheets document
        :return: server response
        """
        self._lgr.info("GoogleUpdate.send_sheets_data()\n")

        response = {'Response': 'None'}
        try:
            assets_body = {
                'valueInputOption': 'USER_ENTERED',
                'data': self._data
            }
            response = self.vals.batchUpdate(spreadsheetId=self.__get_budget_id(), body=assets_body).execute()

            self._lgr.info(F"{response.get('totalUpdatedCells')} cells updated!\n")

        except Exception as ssde:
            msg = repr(ssde)
            self._lgr.error(F"GoogleUpdate.send_sheets_data() Exception: {msg}!")
            response['Response'] = msg

        return response

    # noinspection PyTypeChecker
    def read_sheets_data(self, range_name:str) -> list:
        """
        Get data from my Google sheets document
        :return: server response
        """
        self._lgr.info("GoogleUpdate.read_sheets_data()\n")

        try:
            response = self.vals.get(spreadsheetId = self.__get_budget_id(), range = range_name).execute()
            rows = response.get('values', [])

            self._lgr.info(F"{len(rows)} rows retrieved.\n")

        except Exception as rsde:
            msg = repr(rsde)
            self._lgr.error(F"GoogleUpdate.read_sheets_data() Exception: {msg}!")
            rows = [msg]

        return rows

# END class GoogleUpdate
