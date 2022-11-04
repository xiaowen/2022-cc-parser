import io
import os

from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient

from google.api_core.client_options import ClientOptions
import google.auth
from google.cloud import documentai
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

# Spreadsheet constants
SPREADSHEET_ID = '1hxHkXY0jugNsx35P8i4fQylURkhqgBjoooHG5exSC2Q'
SHEET_NAME_AZURE = 'Azure'
SHEET_NAME_GCLOUD = 'GCloud'

# Google Cloud and Doc AI constants
GCLOUD_PROJECT_ID = 'tensile-howl-307302'
DOCAI_LOCATION = 'us'
DOCAI_PROCESSOR_ID = 'be5c4fa46ae54842'

# Azure constants
AZURE_FORM_RECOGNIZER_ENDPOINT = "https://xiaowenx.cognitiveservices.azure.com/"
AZURE_FORM_RECOGNIZER_KEY = os.environ['AZURE_COGNITIVE_SERVICES_KEY']


def get_sheets_data(sheet_name):
    # Get data from the spreadsheet
    sheet = build('sheets', 'v4').spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=sheet_name+"!A2:D").execute()
    values = result.get('values', [])

    # Reformat to a dictionary with file names as keys
    return dict( (x[0], [i] + x[1:]) for i, x in enumerate(values) )

def append_to_sheet(sheet_name, file_name, stmt_date, balance, note):
    sheet = build('sheets', 'v4').spreadsheets()
    return sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=sheet_name+"!A2:D",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={ 'values': [[file_name, stmt_date, balance, note]]}).execute()

def get_stmt_list():
    folders = [
        '11kCmcnVAwhhfrXwOA8NNeDD_HWZMEQCm', # 2017
        '1oxNbN_dp5uC-CD-PAc2nfS4t_2u3AN70', # 2018
        '1GMErfAByXF0T_miIxGUgJ7UmE1Qz7qSr', # 2019
        '1BTeyaSVZ-tfR6s8gsa9dA9TdyqvAa2UK', # 2020
        '1f-yuZBhcP2B2yw57I2JwNG-2pqZx64gi', # 2021
        '1hiaLLh8Po9uiDHewKhnbryuggRwZiGcv'] # 2022

    stmts = []
    service = build('drive', 'v3')

    for folder_id in folders:
        response = service.files().list(
            q="'%s' in parents and mimeType='application/pdf' and name contains 'citi-20'" % (folder_id),
            spaces='drive').execute()

        stmts.extend([(f.get('name'), f.get('id')) for f in response.get('files', [])])

    return [(n,i) for (n,i) in stmts if n.endswith('-1p.pdf')]

def download_stmt(file_id):
    service = build('drive', 'v3')
    request = service.files().get_media(fileId=file_id)
    file = io.BytesIO()
    downloader = MediaIoBaseDownload(file, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()

    return file.getvalue()

def parse_stmt(cloud, content, file_path=None):
    if cloud == 'gcloud':
        return parse_stmt_gcloud(content, file_path)
    elif cloud == 'azure':
        return parse_stmt_azure(content, file_path)

def parse_stmt_gcloud(image_content, file_path=None):
    opts = ClientOptions(api_endpoint=f"{DOCAI_LOCATION}-documentai.googleapis.com")
    client = documentai.DocumentProcessorServiceClient(client_options=opts)
    name = client.processor_path(GCLOUD_PROJECT_ID, DOCAI_LOCATION, DOCAI_PROCESSOR_ID)

    if file_path:
        with open(file_path, "rb") as image:
            image_content = image.read()

    raw_document = documentai.RawDocument(content=image_content, mime_type='application/pdf')
    request = documentai.ProcessRequest(name=name, raw_document=raw_document)
    result = client.process_document(request=request)
    document = result.document

    stmt_date, balance, note = None, None, ''

    for field in document.pages[0].form_fields:
        field_name = field.field_name.text_anchor.content.strip()
        field_value = field.field_value.text_anchor.content.strip()

        if field_name.startswith('New balance as'):
            note = "Parsed from form fields: '%s' ; '%s' ; " % (field_name, field_value)

            stmt_date = field_name.split()[-1].strip(':')
            balance = field_value.strip(':\n ')

            if len(balance.split()) >= 2:
                stmt_date, balance = balance.split()[-2:]
                stmt_date = stmt_date.strip(':\n ')
                balance = balance.strip(':\n ')

            break

    for table in document.pages[0].tables:
        for row in table.body_rows:
            row_texts = []

            for cell in row.cells:
                if not cell.layout.text_anchor:
                    break

                for segment in cell.layout.text_anchor.text_segments:
                    row_texts.append(document.text[segment.start_index:segment.end_index])

            row_text = ''.join(row_texts)
            if row_text.startswith('New balance as of'):
                note += "Parsed from table: '%s' ; " % (row_texts)

                stmt_date, balance = row_text.strip().split()[-2:]
                stmt_date = stmt_date.strip(':')

                break

    return stmt_date, balance, note

def parse_stmt_azure(image_content, file_path=None):
    if file_path:
        with open(file_path, "rb") as image:
            image_content = image.read()

    document_analysis_client = DocumentAnalysisClient(
        endpoint=AZURE_FORM_RECOGNIZER_ENDPOINT,
        credential=AzureKeyCredential(AZURE_FORM_RECOGNIZER_KEY))

    poller = document_analysis_client.begin_analyze_document("prebuilt-document", document=image_content)
    result = poller.result()

    stmt_date, balance, note = None, None, ''

    # Try to find the info we need in the key/value pairs
    for kv_pair in result.key_value_pairs:
        if kv_pair.key and kv_pair.value:
            if kv_pair.key.content.startswith('New balance as'):
                stmt_date = kv_pair.key.content.split()[-1].strip(':')
                balance = kv_pair.value.content
                note += "KV pairs key '{}': value: '{}' ; ".format(kv_pair.key.content, kv_pair.value.content)

                break

    # If not in key/value pairs, then check the tables data structure
    for tab in result.tables:
        cells = tab.to_dict()['cells']
        if len(cells) >= 5:
            key = cells[3]['content']    
            value = cells[4]['content']

            if key.startswith('New balance as of'):
                stmt_date = key.split()[-1].strip(':')
                balance = value
                note += "Table key '{}': value: '{}'".format(key, value)

                break

    return stmt_date, balance, note

if __name__ == "__main__":
    # Get list of statements from Google Drive
    stmts = get_stmt_list()

    for cloud in ['gcloud', 'azure']:
        print('Working on: ' + cloud)

        # Get the existing info in the spreadsheet
        sheet_name = dict(gcloud=SHEET_NAME_GCLOUD, azure=SHEET_NAME_AZURE)[cloud]
        sheets_data = get_sheets_data(sheet_name)

        for file_name, file_id in stmts:
            if file_name in sheets_data:
                continue # Skip if this file has already been processed

            print('Processing: ' + file_name)

            # Download the file from Google Drive into a buffer
            content = download_stmt(file_id)

            # Parse out the statement date and balance
            stmt_date, balance, note = parse_stmt(cloud, content)

            # Add results to spreadsheet
            append_to_sheet(sheet_name, file_name, stmt_date, balance, note)