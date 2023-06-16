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
SPREADSHEET_IDS = ['1oPDEL4WFHxRQJtelNnODUeP27MVesFvCis-xMyh4E0A', #Giovanna S
                   '16-6u2v-7J1-vsfoUqw32zw95t_wv__-uB87NjKUVX68', #Luciana
                    '1gUsP91E8e0MLIPBi2KcqLtUAKuYU85dZHWTA6CvmrE4', #Giovanna P
                    '11eFXvWal11CDywrCJoSm_ZV885v9nWVosWYcy5yPQpY',#Bianca
                    '1e-840b-da9SYwW1fP7hzvL637osZ8gbEEeztZPUw2yA',#Luana
                    '1SUR0w8DbG1hsUPmKKLGKdyxWzxlTndDudUz-kSH5Nzc',#Luiz
                    '1h0hRZe2kkFT0zypmgH8aJutTuH35LbJRk10WkpCAW-8', #Bete
                    '1MrTWlYR5Q8xLYs_8dNIWKAwJA-3RLcrYJZWAVMepd7Q',#Amanda 
                    '1sYsuXjBbqJJcWYHRxSEeJzChWl-x2IAo0U9v304wEWQ',#Sabatha
                    '1EFaUUo6wv_mziGr1VW6Avc5Iwwbk0cEfqVMndxatBJM',#Andressa
                    '1zhJATymBAGaqmCu07930RlkQjxiwF1vHnJrF0FS0flc', #Karyna
                    '1trPpDGHex-KiccSX0NBa7TgOUWHhk3a2zvICI08Cb_I', #Nicolas
                    '1f-jzoEzF9mgXpv0awBUw_wXFVOgqyA0xv5f387FJjw8'#Julia
                    ] 
RANGE_NAMES = ['FRASE!A1:F', 'FRASE!A1:F', 'FRASE!A1:F', 'FRASE!A1:F', 'FRASE!A1:F' , 'FRASE!A1:F' , 'FRASE!A1:F' , 'FRASE!A1:F' , 'FRASE!A1:F' , 'FRASE!A1:F', 'FRASE!A1:F' ,'FRASE!A1:F', 'FRASE!A1:F']


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