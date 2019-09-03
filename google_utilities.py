##############################################################################################################################
# coding=utf-8
#
# google_utilities.py -- useful constants, functions & classes for Google
#
# some code from Google quickstart examples
#
# Copyright (c) 2019 Mark Sattolo <epistemik@gmail.com>
#
__author__ = 'Mark Sattolo'
__author_email__ = 'epistemik@gmail.com'
__python_version__ = 3.6
__created__ = '2019-04-07'
__updated__ = '2019-08-31'

import sys
sys.path.append("/home/marksa/dev/git/Python/Utilities/")
import os.path as osp
import pickle
from decimal import Decimal
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from python_utilities import *

# sheet names in Budget Quarterly
ALL_INC_SHEET:str    = 'All Inc 1'
ALL_INC_2_SHEET:str  = 'All Inc 2'
NEC_INC_SHEET:str    = 'Nec Inc 1'
NEC_INC_2_SHEET:str  = 'Nec Inc 2'
BAL_1_SHEET:str      = 'Balance 1'
BAL_2_SHEET:str      = 'Balance 2'
QTR_ASTS_SHEET:str   = 'Assets 1'
QTR_ASTS_2_SHEET:str = 'Assets 2'
ML_WORK_SHEET:str    = 'ML Work'
CALCULNS_SHEET:str   = 'Calculations'

# first data row in Budget-qtrly.gsht
BASE_ROW:int = 3


class GoogleUtilities:
    def __init__(self):
        SattoLog.print_text("GoogleUtilities", GREEN)

    my_color = CYAN

    CREDENTIALS_FILE:str = 'secrets/credentials.json'

    SHEETS_RW_SCOPE:list = ['https://www.googleapis.com/auth/spreadsheets']

    SHEETS_EPISTEMIK_RW_TOKEN:dict = {
        'P2' : 'secrets/token.sheets.epistemik.rw.pickle2' ,
        'P3' : 'secrets/token.sheets.epistemik.rw.pickle3' ,
        'P4' : 'secrets/token.sheets.epistemik.rw.pickle4'
    }
    TOKEN:str = SHEETS_EPISTEMIK_RW_TOKEN['P4']

    # Spreadsheet ID
    BUDGET_QTRLY_ID_FILE:str = 'secrets/Budget-qtrly.id'

    @staticmethod
    def fill_cell(sheet:str, col:str, row:int, val, data:list=None) -> list :
        """
        create a dict of update information for one Google Sheets cell and add to the submitted or created list
        :param sheet:  particular sheet in my Google spreadsheet to update
        :param   col:  column to update
        :param   row:  to update
        :param   val:  str OR Decimal: value to fill with
        :param  data:  optional list to append with created dict
        """
        # SattoLog.print_text("GoogleUtilities.fill_cell()", GoogleUtilities.my_color)

        if data is None:
            data = list()

        value = val.to_eng_string() if isinstance(val, Decimal) else val
        cell = {'range': sheet + '!' + col + str(row), 'values': [[value]]}
        SattoLog.print_text("GoogleUtilities.fill_cell() = {}\n".format(cell), GoogleUtilities.my_color)
        data.append(cell)

        return data

    @staticmethod
    def get_budget_id() -> str :
        """
        get the budget id string from the file in the secrets folder
        """
        SattoLog.print_text("GoogleUtilities.get_budget_id()", GoogleUtilities.my_color)

        fp = open(GoogleUtilities.BUDGET_QTRLY_ID_FILE, "r")
        fid = fp.readline().strip()
        SattoLog.print_text("Budget Id = '{}'\n".format(fid), GoogleUtilities.my_color)
        fp.close()
        return fid

    @staticmethod
    def get_credentials() -> pickle :
        """
        get the proper credentials needed to write to the Google spreadsheet
        """
        SattoLog.print_text("GoogleUtilities.get_credentials()", GoogleUtilities.my_color)

        creds = None
        if osp.exists(GoogleUtilities.TOKEN):
            with open(GoogleUtilities.TOKEN, 'rb') as token:
                creds = pickle.load(token)

        # if there are no (valid) credentials available, let the user log in.
        if creds is None or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(GoogleUtilities.CREDENTIALS_FILE,
                                                                 GoogleUtilities.SHEETS_RW_SCOPE)
                creds = flow.run_local_server()
            # save the credentials for the next run
            with open(GoogleUtilities.TOKEN, 'wb') as token:
                pickle.dump(creds, token, pickle.HIGHEST_PROTOCOL)

        return creds

    @staticmethod
    def send_data(mode:str, data:list) -> dict :
        """
        Send the data list to the document
        :param mode: '.?[send][1]'
        :param data: Gnucash data for each needed quarter
        :return: server response
        """
        SattoLog.print_text("GoogleUtilities.send_data({})\n".format(mode), GoogleUtilities.my_color)

        response = None
        try:
            assets_body = {
                'valueInputOption': 'USER_ENTERED',
                'data': data
            }

            if 'send' in mode:
                creds = GoogleUtilities.get_credentials()
                service = build('sheets', 'v4', credentials=creds)
                vals = service.spreadsheets().values()
                response = vals.batchUpdate(spreadsheetId=GoogleUtilities.get_budget_id(), body=assets_body).execute()

                SattoLog.print_text('{} cells updated!\n'.format(response.get('totalUpdatedCells')), GoogleUtilities.my_color)

        except Exception as se:
            SattoLog.print_warning("Exception: {}!".format(se))
            exit(419)

        return response

# END class GoogleUtilities
