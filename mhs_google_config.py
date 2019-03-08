#
# mhs_google_config.py -- constants for Google API's to access files
#
# @author Mark Sattolo <epistemik@gmail.com>
# @revised 2019-03-08
#

CREDENTIALS = 'secrets/credentials.json'

# If modifying these scopes, delete the file 'token.pickle'
DOCS_RO_SCOPE = ['https://www.googleapis.com/auth/documents.readonly']
DOCS_RW_SCOPE = ['https://www.googleapis.com/auth/documents']
DOCS_EPISTEMIK_RW_TOKEN = 'secrets/token.docs.epistemik.rw.pickle'

SHEETS_RO_SCOPE = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SHEETS_RW_SCOPE = ['https://www.googleapis.com/auth/spreadsheets']
SHEETS_SAMPLE_RW_TOKEN    = 'secrets/token.sheets.sample.rw.pickle'
SHEETS_EPISTEMIK_RW_TOKEN = 'secrets/token.sheets.epistemik.rw.pickle'

DRIVE_RO_SCOPE = ['https://www.googleapis.com/auth/drive.readonly']
DRIVE_RW_SCOPE = ['https://www.googleapis.com/auth/drive']


# The ID of a sample document.
SAMPLE_DOCUMENT = '195j9eDD3ccgjQRttHhJPymLJUCOUjs-jmwTrekvdjFE'

# The ID and range of a sample spreadsheet
SAMPLE_SPREADSHEET = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
SAMPLE_RANGE       = 'Class Data!A2:E'

# my documents
CURRENT_READING_DOC = '1VlYk7qu7DFarxYOwK78TeoxAbwA3gCjza751xlDBxzU'
READING_DOC         = '1TeFPDuI1ergAi4RAifoht6XkM-QA-kxdL98eYKEOM6k'

# my spreadsheets
BUDGET_QTRLY_SHEET = '1YbHb7RjZUlA2gyaGDVgRoQYhjs9I8gndKJ0f1Cn-Zr0'
# 'sheet!start_cell:end_cell'
BUDGET_QTRLY_SAMPLE_RANGE = 'All Inc Quarterly!B1:Q10'
