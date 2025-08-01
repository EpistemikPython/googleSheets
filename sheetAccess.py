##############################################################################################################################
# coding=utf-8
#
# sheetAccess.py -- useful constants, functions & classes for access to my Google Sheets
#
# includes some code from Google quickstart examples
#
# Copyright (c) 2025 Mark Sattolo <epistemik@gmail.com>

__author_name__    = "Mark Sattolo"
__author_email__   = "epistemik@gmail.com"
__python_version__ = "3.6+"
__google_api_python_client_version__ = "2.113.0"
__created__ = "2019-04-07"
__updated__ = "2025-07-02"

import threading
from decimal import Decimal
from sys import path
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from typing import Union
path.append("/home/marksa/git/Python/utils")
from mhsUtils import get_current_time, osp, lg, BASE_PYTHON_FOLDER, JSON_LABEL
from mhsLogging import get_simple_logger

# see https://github.com/googleapis/google-api-python-client/issues/299
# use: e.g. build("drive", "v3", http=http, cache_discovery=False)
lg.getLogger("googleapiclient.discovery_cache").setLevel(lg.ERROR)

BASE_ROW:str = "Base Row"
# sheet names in Budget Quarterly
ML_WORK_SHEET:str  = "ML Work"
CALCULNS_SHEET:str = "Calculations"

SECRETS_DIR = osp.join(BASE_PYTHON_FOLDER, "google" + osp.sep + "sheets" + osp.sep + "secrets")
CREDENTIALS_FILE:str = osp.join(SECRETS_DIR, "credentials" + osp.extsep + "json")
SHEETS_RW_SCOPE:list = ['https://www.googleapis.com/auth/spreadsheets']
GGL_SHEETS_TOKEN:str = osp.join(SECRETS_DIR, "token" + osp.extsep + JSON_LABEL)
# Spreadsheet ID
BUDGET_QTRLY_ID_FILE:str = osp.join(SECRETS_DIR, "Budget-qtrly" + osp.extsep + "id")

# sheet names in Budget Quarterly
ALL_INC_SHEET:str    = "All Inc 1"
ALL_INC_2_SHEET:str  = "All Inc 2"
NEC_INC_SHEET:str    = "Nec Inc 1"
NEC_INC_2_SHEET:str  = "Nec Inc 2"
QTR_ASTS_SHEET:str   = "Assets 1"
QTR_ASTS_2_SHEET:str = "Assets 2"
BAL_1_SHEET:str      = "Balance 1"
BAL_2_SHEET:str      = "Balance 2"

FILL_CELL_VAL = Union[str, Decimal]

def get_credentials(lgr:lg.Logger=None) -> Credentials:
    """Get the proper credentials needed to write to the Google spreadsheet."""
    creds = None
    if osp.exists(GGL_SHEETS_TOKEN):
        if lgr: lgr.info(F"osp.exists({GGL_SHEETS_TOKEN})")
        creds = Credentials.from_authorized_user_file(GGL_SHEETS_TOKEN, SHEETS_RW_SCOPE)

    # if there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if lgr: lgr.info("creds is None or not creds.valid")
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            if lgr: lgr.debug("creds.refresh(Request())")
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SHEETS_RW_SCOPE)
            creds = flow.run_local_server()
            if lgr: lgr.debug("creds = flow.run_local_server()")
        # save the credentials for the next run
        with open(GGL_SHEETS_TOKEN, "w") as gst:
            gst.write( creds.to_json() )

    return creds


class MhsSheetAccess:
    """Start a Google session, read/write to my Budget sheet, end the session."""
    # prevent different instances/threads from writing at the same time
    _lock = threading.Lock()

    def __init__(self, p_logger:lg.Logger=None):
        self._lgr = p_logger if p_logger else get_simple_logger(self.__class__.__name__)
        self._data = list()
        self._lgr.info(F"Launch {self.__class__.__name__} instance with lock '{str(self._lock)}' at {get_current_time()}\n")

    def get_data(self) -> list:
        return self._data

    # noinspection PyAttributeOutsideInit
    def begin_session(self):
        # PREVENT starting a separate Session on the Google sheet
        self._lock.acquire()
        self._lgr.info(F"acquired lock at {get_current_time()}")
        creds = get_credentials(self._lgr)
        service = build("sheets", "v4", credentials = creds, cache_discovery = False)
        self.vals = service.spreadsheets().values()

    def end_session(self):
        # RELEASE the Session on the Google sheet
        self._lock.release()
        self._lgr.debug(F"released lock at {get_current_time()}")

    def __get_budget_id(self) -> str:
        """Get the budget id string from the file in the secrets folder."""
        with open(BUDGET_QTRLY_ID_FILE) as bfp:
            fid = bfp.readline().strip()
        self._lgr.debug(F"{get_current_time()} / Budget Id = {fid}\n")
        return fid

    def fill_cell(self, sheet:str, col:str, row:int, val:FILL_CELL_VAL):
        """
        CREATE the information to update a Google Sheets cell and add to the data list
        :param sheet: particular sheet in my Google spreadsheet to update
        :param   col: column to update
        :param   row: to update
        :param   val: str OR Decimal: value to fill with
        """
        self._lgr.debug( get_current_time() )
        value = val.to_eng_string() if isinstance(val, Decimal) else val
        cell = {"range": sheet + '!' + col + str(row), "values": [[value]]}
        self._lgr.debug(F"fill_cell() = {cell}\n")
        self._data.append(cell)

    def send_sheets_data(self) -> dict:
        """
        SEND the data list to my Google sheets document
        :return: server response
        """
        self._lgr.debug( get_current_time() )
        if not self.vals:
            msg = "No Session started!"
            self._lgr.exception(msg)
            return {"PROBLEM": msg}

        try:
            assets_body = {
                "valueInputOption": "USER_ENTERED",
                "data": self._data
            }
            response = self.vals.batchUpdate(spreadsheetId=self.__get_budget_id(), body=assets_body).execute()
            self._lgr.info(F"{response.get('totalUpdatedCells')} cells updated.")
        except Exception as ssde:
            self._lgr.error(ssde)
            raise ssde

        return response

    def read_sheets_data(self, range_name:str) -> list:
        """
        READ data from my Google sheets document
        :return: server response
        """
        self._lgr.debug( get_current_time() )
        if not self.vals:
            msg = "No Session started!"
            self._lgr.exception(msg)
            return [msg]

        try:
            response = self.vals.get(spreadsheetId = self.__get_budget_id(), range = range_name).execute()
            rows = response.get("values", [])
            self._lgr.info(F"{len(rows)} rows retrieved.")
        except Exception as rsde:
            self._lgr.error(rsde)
            raise rsde

        return rows

    def test_read(self, range_name:str) -> list:
        self.begin_session()
        result = self.read_sheets_data(range_name)
        self.end_session()
        return result
# END class MhsSheetAccess


def mhs_class_test():
    mhs = MhsSheetAccess()
    test_range = "'Record'!A1"
    response = mhs.test_read(test_range)
    print( repr(response) )


if __name__ == "__main__":
    mhs_class_test()
    exit()
