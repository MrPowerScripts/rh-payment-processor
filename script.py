from __future__ import print_function
import pickle
import os.path
import re
import sys
import csv
from datetime import date
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import xlwt 
from xlwt import Workbook
# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.


def main(b):
    SAMPLE_SPREADSHEET_ID = '1M-70XYb8Nw2eaeNTJiy-5TWzgtFCrhWid4973Nb5MJ4'
    SAMPLE_RANGE_NAME = b

    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
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
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        return values
def DateFormat(a):
    return [a[2],a[0],a[1]]
def check(a,b):
    for j in range(3):
        if int(b[j])<int(a[j]):
            return True
        elif int(b[j])>int(a[j]):
            return False
    return True
def report(email,l,pa,pe,po,Date):
    for j in l:
        if check(DateFormat(re.split('-|/',str(j[-1]))),Date):
            if j[0].upper() in email:
                email[j[0].upper()][1]+=pa
            if j[1].upper() in email:
                email[j[1].upper()][1]+=pe
            if j[2].upper() in email:  
                email[j[2].upper()][1]+=po
    return email

if __name__ == '__main__':
    count,row=1,1
    wb = Workbook() 
    sheet1 = wb.add_sheet('Sheet 1') 
    sheet1.write(0, 0, 'Email Id')
    sheet1.write(0, 1, 'Amount')
    sheet1.write(0, 2, 'Currency')
        
    if not sys.argv[1]:
      Date=re.split('-|/',input("Enter Date as MM-DD-YYYY :  "))
    else:
      Date=re.split('-|/',sys.argv[1])
    Date=DateFormat(Date)
    today,email= date.today(),{}
    l=main('Articles!D2:G')
    e=main('People!H:H')
    for j in main('People!A2:B'):
        if j[0] == '' or j[1]== '':
            count+=1
            continue
        email[(j[0]+', '+j[1]).upper()]=["None" if count>=len(e) else e[count],0]
        count+=1
    #print(l)
    #print(email)
    p_Author,p_Editor,p_operation=map(int,main('Rates!A2:C2')[0])
    new_score = report(email,l,p_Author,p_Editor,p_operation,Date)
    csvRows = []
    resultFile = open("payments.csv",'w')
    writer = csv.writer(resultFile)
    for j in new_score:
        if new_score[j][1]>0:
            sheet1.write(row,0,new_score[j][0][0])
            sheet1.write(row,1,new_score[j][1])
            sheet1.write(row,2,"USD")
            row+=1
            print(new_score[j][0][0],",",new_score[j][1],",","USD")
            writer.writerow([new_score[j][0][0],new_score[j][1],"USD"])
    # wb.save('Report.csv')
