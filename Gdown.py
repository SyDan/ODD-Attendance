#!/usr/bin/env python

from __future__ import print_function
import sys
import io
import time
#from io import FileIO
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/drive'
store = file.Storage('storage.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
    creds = tools.run_flow(flow, store, flags) \
            if flags else tools.run(flow, store)
DRIVE = build('drive', 'v3', http=creds.authorize(Http()))

results = DRIVE.files().list(pageSize=10,fields="nextPageToken, files(id, name)").execute()
items = results.get('files', [])
if not items:
    print('No files found in your drive.')
else:
    print('The list of Files in your GOOGLE DRIVE are:')
    i=0
    for item in items:
        print('{0}.\t{1}'.format(i, item['name']))
        i+=1
num = input('Enter the serial no. of File you want to download: ')
while num>=i:
    print('Invalid Serial No. Please Input Again!!')
    num = input('Enter the serial no. of File you want to download: ')
else:
    MIMETYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    data = DRIVE.files().get_media(fileId=items[num]['id'])
    fh = io.FileIO(items[num]['name'], 'wb')
    downloader = MediaIoBaseDownload(fh, data, chunksize=40000)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print ('Download %d%%.' % int(status.progress() * 100))
    print('Download Complete')
time.sleep(2)
sys.exit()