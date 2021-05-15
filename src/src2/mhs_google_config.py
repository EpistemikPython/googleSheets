#
# mhs_google_config.py -- constants for Google API's to access files
#
# @author Google
# @revised Mark Sattolo <epistemik@gmail.com>
# @version Python3.6
# @created 2019-02-23
# @updated 2019-03-31


CREDENTIALS = 'secrets/credentials.json'

# If modifying these scopes, delete the file 'token.pickle'
DOCS_RO_SCOPE = ['https://www.googleapis.com/auth/documents.readonly']
DOCS_RW_SCOPE = ['https://www.googleapis.com/auth/documents']

SHEETS_RO_SCOPE = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SHEETS_RW_SCOPE = ['https://www.googleapis.com/auth/spreadsheets']

DRIVE_RO_SCOPE = ['https://www.googleapis.com/auth/drive.readonly']
DRIVE_RW_SCOPE = ['https://www.googleapis.com/auth/drive']

DOCS_EPISTEMIK_RW_TOKEN = {
    'P2' : 'secrets/token.docs.epistemik.rw.pickle2' ,
    'P3' : 'secrets/token.docs.epistemik.rw.pickle3' ,
    'P4' : 'secrets/token.docs.epistemik.rw.pickle4'
}

SHEETS_SAMPLE_RW_TOKEN_P2 = 'secrets/token.sheets.sample.rw.pickle2'

SHEETS_EPISTEMIK_RW_TOKEN = {
    'P2' : 'secrets/token.sheets.epistemik.rw.pickle2' ,
    'P3' : 'secrets/token.sheets.epistemik.rw.pickle3' ,
    'P4' : 'secrets/token.sheets.epistemik.rw.pickle4'
}

# The ID of a sample document.
SAMPLE_DOCUMENT = '195j9eDD3ccgjQRttHhJPymLJUCOUjs-jmwTrekvdjFE'

# The ID and range of a sample spreadsheet
SAMPLE_SPREADSHEET = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
SAMPLE_RANGE       = 'Class Data!A2:E'

# my documents
CURRENT_READING_DOC = '1VlYk7qu7DFarxYOwK78TeoxAbwA3gCjza751xlDBxzU'
READING_DOC         = '1TeFPDuI1ergAi4RAifoht6XkM-QA-kxdL98eYKEOM6k'

# my spreadsheets
BUDGET_QTRLY_SPRD_SHEET = '1YbHb7RjZUlA2gyaGDVgRoQYhjs9I8gndKJ0f1Cn-Zr0'
# sheet ids in Budget Quarterly
BUDQTR_ALL_INC_SHEET  = '1581653901'
BUDQTR_ALL_INC_PRAC_SHEET  = '1684660496'
BUDQTR_NEC_INC_SHEET  = '352534630'
BUDQTR_NEC_INC_PRAC_SHEET  = '317111001'
BUDQTR_BALANCE_SHEET  = '1092295261'
BUDQTR_QTR_ASTS_SHEET = '1868004173'
BUDQTR_ML_WORK_SHEET  = '1366666149'
BUDQTR_CALCULNS_SHEET = '1533312865'

# 'sheet!start_cell:end_cell'
BUDGET_QTRLY_READ_RANGE = 'All Inc Quarterly!B1:Q10'
BUDGET_QTRLY_WRITE_RANGE = 'Calculations!A56'
