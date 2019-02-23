#
# createGnucashTxs.py -- parse a Monarch record, possibly from a json file,
#                        create Gnucash transactions from the data
#                        and write to a Gnucash file
#
# Copyright (c) 2018, 2019 Mark Sattolo <epistemik@gmail.com>
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License as
# published by the Free Software Foundation; either version 2 of
# the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# @author Mark Sattolo <epistemik@gmail.com>
#

# If modifying these scopes, delete the file token.pickle
DOCS_RO_SCOPE = ['https://www.googleapis.com/auth/documents.readonly']
DOCS_RW_SCOPE = ['https://www.googleapis.com/auth/documents']

SHEETS_RO_SCOPE = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SHEETS_RW_SCOPE = ['https://www.googleapis.com/auth/spreadsheets']

DRIVE_RO_SCOPE = ['https://www.googleapis.com/auth/drive.readonly']
DRIVE_RW_SCOPE = ['https://www.googleapis.com/auth/drive']

# The ID of a sample document.
SAMPLE_DOCUMENT = '195j9eDD3ccgjQRttHhJPymLJUCOUjs-jmwTrekvdjFE'

CURRENT_READING_DOC = '1VlYk7qu7DFarxYOwK78TeoxAbwA3gCjza751xlDBxzU'
BUDGET_QTRLY_SHEET = '1YbHb7RjZUlA2gyaGDVgRoQYhjs9I8gndKJ0f1Cn-Zr0'
