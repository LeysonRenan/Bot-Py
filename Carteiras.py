from __future__ import print_function
import os.path
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Se for modificar esses escopos, exclua o arquivo token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# IDs e intervalos das planilhas de exemplo.
SPREADSHEET_IDS = ['ID_DA_PLANILHA', #Giovanna S
                   ' ID_DA_PLANILHA ', #Luciana
                    ' ID_DA_PLANILHA ', #Giovanna P
                    ' ID_DA_PLANILHA ',#Bianca
                    ' ID_DA_PLANILHA ',#Luana
                                 ] 
RANGE_NAMES = ['FRASE!A1:F', 'FRASE!A1:F', 'FRASE!A1:F', 'FRASE!A1:F', 'FRASE!A1:F']


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        for i in range(len(SPREADSHEET_IDS)):
            spreadsheet_id = SPREADSHEET_IDS[i]
            range_name = RANGE_NAMES[i]

            sheet = service.spreadsheets()

            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
            df = pd.read_excel(r'.\\Att2.xlsx')
            data = df.values.tolist()

            sheet.values().update(spreadsheetId=spreadsheet_id,
                                  range=range_name,
                                  valueInputOption='RAW',
                                  body={'values': data}).execute()

            print('Dados copiados e colados com sucesso na planilha com o ID:', spreadsheet_id)

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
