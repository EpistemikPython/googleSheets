##############################################################################################################################
# coding=utf-8
#
# google_utilities.py -- useful constants, functions & classes for Google
#
# some code from Google quickstart examples
#
# Copyright (c) 2019 Mark Sattolo <epistemik@gmail.com>
#
from typing import Union

__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__ = 3.6
__created__ = '2019-04-07'
__updated__ = '2019-09-10'

from sys import path
path.append("/home/marksa/dev/git/Python/Utilities/")
import os.path as osp
import pickle
from decimal import Decimal
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from python_utilities import SattoLog

# sheet names in Budget Quarterly
ML_WORK_SHEET:str  = 'ML Work'
CALCULNS_SHEET:str = 'Calculations'

# first data row in Budget-qtrly.gsht
BASE_ROW:int = 3

CREDENTIALS_FILE:str = 'secrets/credentials.json'

SHEETS_RW_SCOPE:list = ['https://www.googleapis.com/auth/spreadsheets']

SHEETS_EPISTEMIK_RW_TOKEN:dict = {
    'P2' : 'secrets/token.sheets.epistemik.rw.pickle2' ,
    'P3' : 'secrets/token.sheets.epistemik.rw.pickle3' ,
    'P4' : 'secrets/token.sheets.epistemik.rw.pickle4'
}
GGL_SHEETS_TOKEN:str = SHEETS_EPISTEMIK_RW_TOKEN['P4']

# Spreadsheet ID
BUDGET_QTRLY_ID_FILE:str = 'secrets/Budget-qtrly.id'

FILL_CELL_VAL = Union[str, Decimal]


class GoogleUpdate:
    def __init__(self, p_logger:SattoLog=None):
        self.logger = p_logger
        self.data = list()

    def get_data(self) -> list :
        return self.data

    def __log(self, msg:str):
        if self.logger:
            self.logger.print_info(msg)

    def fill_cell(self, sheet:str, col:str, row:int, val:FILL_CELL_VAL):
        """
        create a dict of update information for one Google Sheets cell and add to the submitted or created list
        :param   sheet: particular sheet in my Google spreadsheet to update
        :param     col: column to update
        :param     row: to update
        :param     val: str OR Decimal: value to fill with
        """
        self.__log("GoogleUpdate.fill_cell()")

        value = val.to_eng_string() if isinstance(val, Decimal) else val
        cell = {'range': sheet + '!' + col + str(row), 'values': [[value]]}
        self.__log("fill_cell() = {}\n".format(cell))
        self.data.append(cell)

    def __get_budget_id(self) -> str :
        """
        get the budget id string from the file in the secrets folder
        """
        self.__log("google_utilities.__get_budget_id()")

        fp = open(BUDGET_QTRLY_ID_FILE, "r")
        fid = fp.readline().strip()
        self.__log("GGLU.__get_budget_id(): Budget Id = '{}'\n".format(fid))
        fp.close()

        return fid

    def __get_credentials(self) -> pickle :
        """
        get the proper credentials needed to write to the Google spreadsheet
        """
        self.__log("google_utilities.__get_credentials()")

        creds = None
        if osp.exists(GGL_SHEETS_TOKEN):
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
                pickle.dump(creds, token, pickle.HIGHEST_PROTOCOL)

        return creds

    def send_sheets_data(self) -> dict :
        """
        Send the data list to my Google sheets document
        :return: server response
        """
        self.__log("google_utilities.send_sheets_data()\n")

        response = {'Response': 'None'}
        try:
            assets_body = {
                'valueInputOption': 'USER_ENTERED',
                'data': self.data
            }

            creds = self.__get_credentials()
            service = build('sheets', 'v4', credentials=creds)
            vals = service.spreadsheets().values()
            response = vals.batchUpdate(spreadsheetId=self.__get_budget_id(), body=assets_body).execute()

            self.__log('{} cells updated!\n'.format(response.get('totalUpdatedCells')))

        except Exception as sde:
            msg = repr(sde)
            SattoLog.print_warning("Exception: {}!".format(msg))
            response['Response'] = msg

        return response

# END class GoogleUpdate
