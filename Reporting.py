from __future__ import print_function
import io
from urllib.parse import parse_qsl
from googleapiclient.http import MediaIoBaseDownload
import requests
import json
import csv
import requests.auth
import time
import sys
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email
import smtplib
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from itertools import combinations_with_replacement as cwr
import string

today = str(date.today().strftime('%m-%d-%Y') )
alphabet = string.ascii_lowercase
length = 2
sizer = ["".join(comb) for comb in cwr(alphabet, length)]

tag = []
with open('ReportingConfig.json') as config_file:
    Config = json.load(config_file)

def getData(Table,RID,Num=10000):

    url = 'https://api.quickbase.com/v1/reports/'+str(RID)+'/run?tableId='+Table+'&top=10000'

    payload = {}

    headers = Config['Config']['header']

    response = json.loads(requests.request("POST", url, headers=headers, json = payload).text)
    Data = []
    headersrow=[]
    IDS = []
    for x in response['fields']:
        headersrow.append(x['label'])
        IDS.append(int(x['id']))

    Data.append(headersrow)

    print(IDS)
    print(len(response['data']))
    for x in response['data']:
        Row=[]
        for z in headersrow:
            Row.append("")
        
        
        for y in x:

            stry = int(y)
            idex = IDS.index(stry)
            Row[idex]=(x[y]['value'])

        Data.append(Row)
    return Data

def send(subject,body='',attachment='',receiver = Config['Email']['receiver']):
    sender = Config['Email']['sender']
    server = Config['Email']['server']
    port = Config['Email']['port']
    password = Config['Email']['password']

    message=MIMEMultipart()
    message['From']=Config['Email']['From']
    recipients = receiver
    message['To'] = ", ".join(recipients)
    print(recipients)
    message['Subject']=subject
    message['reply-to'] = Config['Email']['ReplyTo']
    message.attach(MIMEText(body,"plain"))

    
    if not attachment=='':
        with open(attachment,'rb') as f:
            part=email.mime.base.MIMEBase('application','octet-stream')
            part.set_payload(f.read())
        email.encoders.encode_base64(part)
        part.add_header("Content-Disposition","attachment; filename="+attachment)
        message.attach(part)
    with smtplib.SMTP_SSL(server,port) as s:
        s.login(sender, password)
        s.sendmail(sender,recipients,message.as_string())
    print('Sent')

def googleSheetAuth():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
            
    service = build('sheets', 'v4', credentials=creds)

    return service

def makeSheet(name):
    service= googleSheetAuth()
    spreadsheet = {
        'properties': {
            'title': name
        }
    }
    spreadsheet = service.spreadsheets().create(body=spreadsheet,fields='spreadsheetId').execute()
    SPREADSHEET_ID = spreadsheet.get('spreadsheetId')

    return SPREADSHEET_ID

def googleSheetWrite(values,spred_ID):
    service= googleSheetAuth()
    sheet = service.spreadsheets()

    body = {
        'values': values
    }
    range = 'Sheet1!A:'+sizer[len(values[0])+1]
    result = service.spreadsheets().values().update(
        spreadsheetId=spred_ID, range=range,
        valueInputOption='USER_ENTERED', body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))
    time.sleep(10)

    print(range)
    result = sheet.values().get(spreadsheetId=spred_ID,
                                range=range).execute()

def googleDriveAuth():
    SCOPES = ['https://www.googleapis.com/auth/drive']
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('Drive.pickle'):
        with open('Drive.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('Drive.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    return service
                              
def googleMoveFile(file_id,folder_id):

    drive_service = googleDriveAuth()
    # Retrieve the existing parents to remove
    file = drive_service.files().get(fileId=file_id,
                                    fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    # Move the file to the new folder
    file = drive_service.files().update(fileId=file_id,
                                        addParents=folder_id,
                                        removeParents=previous_parents,
                                        fields='id, parents',
                                        supportsAllDrives=True).execute()

def googleDownloadFile(file_id,filename):
    drive_service = googleDriveAuth()
    request = drive_service.files().export_media(fileId=file_id,mimeType='text/csv')
    fh = io.FileIO(filename, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))


def main():
    Name = sys.argv[1]+"-"+today
    
    Report=sys.argv[2]
    Table=sys.argv[3]


    sheet = makeSheet(Name)
    data = getData(Table,Report)
    if len(data) > 1:
        googleSheetWrite(data,sheet)
        googleMoveFile(sheet,Config['Config']['GoogleFolder'])
        file = Name+'.csv'
        googleDownloadFile(sheet,file)
        send(Name,body='Please see Attached File.\n ',attachment=file,receiver=sys.argv[4:])
        os.remove(file)
    else:
        send(Name,body='No Devices to report.\n ',receiver=sys.argv[4:])

if __name__ == '__main__':
    main()
    

    
