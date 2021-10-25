import os
from pathlib import Path 
from dotenv import load_dotenv
from slack import WebClient
from slack.errors import SlackApiError
from flask import Flask, request, Response
from slackeventsapi import SlackEventAdapter
import time
import gspread
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.service_account import ServiceAccountCredentials

env_path = Path('.') / '.env'
load_dotenv(dotenv_path=env_path)

app = Flask(__name__)

SLACK_SIGNING_SECRET = os.environ["SIGNING_SECRET"]
slack_events_adapter = SlackEventAdapter(SLACK_SIGNING_SECRET,'/slack/events', app)

sclient = WebClient(token=os.environ['SLACK_TOKEN'])
BOT_ID = sclient.api_call("auth.test")['user_id']

scope = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/drive.file']
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
gclient = gspread.authorize(creds)

# access the drive
gauth = GoogleAuth()
# gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)

# upload dataset
def file_upload():
    f = drive.CreateFile(metadata={"title": "campaignTable"})
    f.SetContentFile('dataset.xlsx')
    f.Upload({'convert': True})
    f.InsertPermission({'type':'user','value': 'goraman-23@campaignautomate.iam.gserviceaccount.com','role': 'writer'})
    return True

def budget(context):
    a = file_upload()
    if a:
        # values of dataset
        sheet = gclient.open("campaignTable").sheet1
        # data = sheet.get_all_records()
        rows = sheet.get_values()
        campaing = sheet.col_values(1)
        headers = sheet.row_values(1)
        C = sheet.col_values(3)
        E = sheet.col_values(5)
        # print(data)
        # print(rows)
        # print(campaing)
    # row by row create files
    for i in range(1,len(campaing)):
        folder_name = campaing[i]
        folder = drive.CreateFile({'title' : folder_name, 'mimeType' : 'application/vnd.google-apps.spreadsheet'})
        folder.Upload()
        folder.InsertPermission({'type':'user','value': 'goraman-23@campaignautomate.iam.gserviceaccount.com','role': 'writer'})
        sh = gclient.open(folder_name)
        time.sleep(1)
        worksheet = sh.sheet1
        # calculates total budget
        for j in range(1,len(rows[i])):
            cell_head = worksheet.update_cell(1,j, headers[j])
            cell_obj = worksheet.update_cell(2,j, rows[i][j])
            if (j == 5):
                # print(C[i])
                # print(E[i])
                a =int(C[i])
                b = int(E[i].replace('.', ''))
                c = f'{a*b:,}'
                total_budget = c.replace(',','.')
                budget_head = worksheet.update_cell(1,6, 'Total Budget')
                cell_newobj = worksheet.update_cell(2,6, total_budget)
                if context == folder_name:
                    calculated = gclient.open(context).sheet1.cell(2,6).value
                    # print(calculated)
        return calculated

@app.route('/calculate_budget', methods=['POST'])
def calculator():
    data = request.form
    # print(data)
    # print(data['command'])
    # print(data['text'])
    command = data['command']
    campaing_name = data['text']
    if command == "/calculate_budget":
        if len(campaing_name) == 1:
            resp = budget(campaing_name)
            sclient.chat_postMessage(channel='#notifications', text=resp)
    return Response(), 200

if __name__ == "__main__":
    app.run(debug=True)
