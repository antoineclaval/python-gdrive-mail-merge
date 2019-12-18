# -*- coding: utf-8 -*-
#
# Copyright ©2018-2019 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at apache.org/licenses/LICENSE-2.0.
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""
docs-mail-merge.py (Python 2.x or 3.x)
Google Docs (REST) API mail-merge sample app
"""
# [START mail_merge_python]
from __future__ import print_function
import time , sys , locale

from googleapiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools

# Localized Date Settings
locale.setlocale(locale.LC_TIME, "fr_FR")

# Fill-in IDs of your Docs template & any Sheets data source
DOCS_FILE_ID = '1UlFsfrDn4QXsdcsTrMZkmz8SXD5tvTVioHWQMBaNtQA'
SHEETS_FILE_ID = '1vvktOMS-GYFh9KFRo-0ZTAm47PP9t9FF45RAVvy8zac'

# authorization constants
CLIENT_ID_FILE = 'credentials.json'
TOKEN_STORE_FILE = 'token.json'
SCOPES = (  # iterable or space-delimited string
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
)

# application constants
SOURCES = ('text', 'sheets')
SOURCE = 'sheets' # Choose one of the data SOURCES
COLUMNS = ['to_name', 'to_title', 'to_company', 'to_address']
C2 = ['DATE SOUMISSION', 'REPONSE_FESTIVAL']
# MOVIE_NAME TARGET_MONTH TARGET_YEAR INSCRIPTIONS_LIST SELECTIONS_LIST REJECTIONS_LIST PROJECTIONS_LIST
# 'Nom du festival'  Pays A D 0 3

def get_http_client():
    """Uses project credentials in CLIENT_ID_FILE along with requested OAuth2
        scopes for authorization, and caches API tokens in TOKEN_STORE_FILE.
    """
    store = file.Storage(TOKEN_STORE_FILE)
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_ID_FILE, SCOPES)
        creds = tools.run_flow(flow, store)
    return creds.authorize(Http())

# service endpoints to Google APIs
HTTP = get_http_client()
DRIVE = discovery.build('drive', 'v3', http=HTTP)
DOCS = discovery.build('docs', 'v1', http=HTTP)
SHEETS = discovery.build('sheets', 'v4', http=HTTP)

def get_data(source):
    """Gets mail merge data from chosen data source.
    """
    if source not in {'sheets', 'text'}:
        raise ValueError('ERROR: unsupported source %r; choose from %r' % (
            source, SOURCES))
    return SAFE_DISPATCH[source]()

def _get_sheets_data(service=SHEETS, range="Sheet1"):
    """(private) Returns data from Google Sheets source. It gets all rows of
        'Sheet1' (the default Sheet in a new spreadsheet), but drops the first
        (header) row. Use any desired data range (in standard A1 notation).
    """
    return service.spreadsheets().values().get(spreadsheetId=SHEETS_FILE_ID,
            range=range).execute().get('values')[1:] # skip header row

# data source dispatch table [better alternative vs . eval()] # will call _get_SOMETHING_data funct
SAFE_DISPATCH = {k: globals().get('_get_%s_data' % k) for k in SOURCES}

def _copy_template(tmpl_id, source, service):
    """(private) Copies letter template document using Drive API then
        returns file ID of (new) copy.
    """
    body = {'name': F'{targetMovie} - {targetMonth} {targetYear}'}

    return service.files().copy(body=body, fileId=tmpl_id,
            fields='id').execute().get('id')

def merge_template(tmpl_id, source, service):
    """Copies template document and merges data into new copy, then
        returns the copy file ID.
    """
    # copy template and set context data struct for merging template values
    copy_id = _copy_template(tmpl_id, source, service)
    context = merge.iteritems() if hasattr({}, 'iteritems') else merge.items()

    print ("--- context ---")
    print ( context)

    # "search & replace" API requests for mail merge substitutions
    reqs = [{'replaceAllText': {
                'containsText': {
                    'text': '{{%s}}' % key.upper(), # {{VARS}} are uppercase
                    'matchCase': True,
                },
                'replaceText': value,
            }} for key, value in context]


    print ( "request :")
    print (reqs)
    # send requests to Docs API to do actual merge
    DOCS.documents().batchUpdate(body={'requests': reqs}, documentId=copy_id, fields='').execute()
    return copy_id


if __name__ == '__main__':

    if len(sys.argv) != 4:
        print("usage : MOVIE MONTH YEAR")
        raise ValueError("Always 3 params : MOVIE MONTH YEAR")

    targetMovie = sys.argv[1]

    if sys.argv[2]  not in {'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'}:
        raise ValueError("Month must be within : Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre")
    
    targetMonth = sys.argv[2]

    targetYear = sys.argv[3]

    print("Generating Montly report for", targetMovie, targetMonth, targetYear)

    # fill-in your data to merge into document template variables
    merge = {

        'MOVIE_NAME' : targetMovie,
        # movie independant data 
        'CURRENT_DATE': time.strftime('%d %B %Y'),
        'TARGET_MONTH':targetMonth,
        'TARGET_YEAR':targetYear,
        # - - - - - - - - - - - - - - - - - - - - - - - - - -
        # Report data, gathered in MOVIE_NAME Sheet
        'INSCRIPTIONS_LIST': None, # toute les lignes avec TARGET_MONTH = A
        'SELECTIONS_LIST': None,  # A + filter "selectionee"
        'REJECTIONS_LIST': None, # A + filter "rejections"
        'PROJECTIONS_LIST': None,
        # - - - - - - - - - - - - - - - - - - - - - - - - - -

    }

    # get row data, then loop through & process each form letter
    dataDict = _get_sheets_data(SHEETS, targetMovie) # get data from data source for targetMovie
  #  print(dataDict)
    # print(dataDict[0][0])
    # print(dataDict[3])

    inscriptionsListString, refusalListString, acceptanceListString , projectionListString  = "" , "", "" , ""
    currentMonthLine = []


    for i in dataDict:
        print(i)
        if  i :
            if  targetMonth in i[1]:
                inscriptionsListString += i[0] + " (" + i[4] + ")\n"
                currentMonthLine.append(i)
            if i[3] and targetMonth in i[3]:
                if( "REFUSÉ" == i[2]):
                    refusalListString += "- " + (i[0] + " (" + i[4] + ")\n")
                elif ( "SÉLECTIONNÉ" ==  i[2]):
                    acceptanceListString += "- " + (i[0] + " (" + i[4] + ")\n")



    print ("Inscription : ", len(currentMonthLine))

    if ( not inscriptionsListString ) :
        inscriptionsListString = "Pas encore"

    if ( not refusalListString  ) :
        refusalListString = "Pas encore"

    if ( not acceptanceListString ) :
        acceptanceListString = "Pas encore"


    merge.update(dict({"INSCRIPTIONS_LIST" : inscriptionsListString}))
    merge.update(dict({"SELECTIONS_LIST" : acceptanceListString}))
    merge.update(dict({"REJECTIONS_LIST" : refusalListString}))
    merge.update(dict({"PROJECTIONS_LIST" : projectionListString}))
    

    print('Merged letter %d: docs.google.com/document/d/%s/edit' % (1, merge_template(DOCS_FILE_ID, SOURCE, DRIVE)))
# [END mail_merge_python]
