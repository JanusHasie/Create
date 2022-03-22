# NB Who runs this program
curruser = 'Janus_Haasbroek'

# Imports
import os
import win32com.client
from datetime import datetime,timedelta
import time, re, shutil, os, pandas as pd
from sqlalchemy import create_engine, text
import datetime 
#import time, re, shutil
# import time
# import logging
# import smtplib
# from sys import exit
import psutil
import subprocess

# Mail Details
MyEmail = 'Murray Smith'
FinishedFolder = 'Data Archive'
SenderName = 'bipublisher-report@oracle.com'
Subject = 'Unrestricted ROMT SBC Report'
FileExtension = 'xlsx'
SaveFolder = 'C:\\Users\\haasbroekj\\Documents\\Fourier\\Sasol\\Data Import'
Maillist = ('haasbroekj@fourier.co.za', 'krugerj@fourier', 'voudtshoornm@fourier.co.za')
print('Mail Details Updated')

# SQL Connection
engine = create_engine("mssql+pymssql://{user}:{pw}@197.189.232.50/{db}"
                       .format(user="sa", pw="NewFAsys098!", db="Sasol_Transport_DB"))
print('SQL Engine Created')

# Email Creation
def send_notification_success():
    #result = runSP(1)
    outlook_mail = win32com.client.Dispatch('outlook.application')
    mail_ = outlook_mail.CreateItem(0)
    #for person in Maillist :
    mail_.To = 'haasbroekj@fourier.co.za'
    mail_.Subject = 'Data Update - Pass'
    mail_.body = 'The database was successfully updated. Stored procedure run code: '#+result
    mail_.Send()


def send_notification_failure():
    #result = runSP(2)
    outlook_mail = win32com.client.Dispatch('outlook.application')
    mail_ = outlook_mail.CreateItem(0)
    for person in Maillist :
        mail_.To = person
        mail_.Subject = 'Data Update - Fail'
        mail_.body = 'The database failed to update. Stored procedure run code: '#+result
        mail_.Send()


def open_outlook():
    try:
        subprocess.call(['C:\\Program Files\\Microsoft Office\\root\\Office16\\Outlook.exe'])
        os.system("C:\\Program Files\\Microsoft Office\\root\\Office16\\Outlook.exe")
        print('Outlook open')
    except:
        print("Outlook didn't open successfully")

'''def runSP(PF) :
    if PF == 1 :
        result = engine.execute(text("EXEC [dbo].[FY22TESTPROC] {User}, {Status}".format(User = curruser, Status = 'Pass')))
        result.fetchall()
        engine.execute(text("EXEC [dbo].[FY22TESTPROC]"))
        print(result)
    elif PF == 2 :
        result = engine.execute(text("EXEC [dbo].[AnalyseFY22] {User}, {Status}".format(User = curruser, Status = 'Fail')))
    return result'''

for item in psutil.pids():
    p = psutil.Process(item)
    if p.name() == "OUTLOOK.EXE":
        flag = 1
        break
    else:
        flag = 0
print(flag)

# Function
def _right(s, amount):
    return s[-amount:]


outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")

Inbox = outlook.Folders[MyEmail].Folders['Inbox']

ProcessedFolder = outlook.Folders[MyEmail].Folders[FinishedFolder]

messages = Inbox.Items
received_dt = datetime.datetime.now() - timedelta(minutes=600)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
messages = messages.Restrict(f'[SenderName] = {SenderName}')
# messages = messages.Restrict("[Subject] = 'Subject'")
print('Mail Ready')

for mail in messages:
    for att in mail.Attachments:
        # if str.lower(_right(att.FileName, len(FileExtension))) == str.lower(FileExtension):
        temp_filename = os.path.normpath(os.path.join(SaveFolder, f'{att.FileName}'))
        print(temp_filename)
        print(att.FileName)
        try:
            att.SaveAsFile(temp_filename)
            print('File Successfully Saved [{}]'.format(temp_filename))
            b_processed = True
            df = pd.read_excel('C:\\Users\\haasbroekj\\Documents\\Fourier\\Sasol\\Data Import\\'+att.FileName)
            df = df[9:]
            df = df[:-4]
            df.columns = df.iloc[0]
            df = df.drop(df.index[0])
            print('File Read')
            df.to_sql('BWLData22', con=engine, if_exists='append', chunksize=1000, index=False)
            print('File Uploaded')
            os.rename("C:\\Users\\haasbroekj\\Documents\\Fourier\\Sasol\\Data Import\\"+att.FileName,
                       "C:\\Users\\haasbroekj\\Documents\\Fourier\\Sasol\\Data Import\\Archive\\"+str(datetime.date.today())+" "+att.FileName)

            
            if (flag == 1):
                print("Send success notification")
                send_notification_success()
            else:
                open_outlook()
                print("Send success notification")
                send_notification_success()
                #engine.execute(text('''EXEC [dbo].[ImportUpdateLog] @Param1 = ?, @Status = ?''', (curruser, 'Pass')).execution_options(autocommit=True))
            print('Daily Update Completed')
        except Exception as e:
            print(str(e) + ' | File NOT saved [{}]'.format(temp_filename))
            if (flag == 1):
                print("Send failure notification")
                send_notification_failure()
                #engine.execute(text('''EXEC [dbo].[ImportUpdateLog] @Param1 = ?, @Status = ?''', (curruser, 'Fail')).execution_options(autocommit=True))
            else:
                open_outlook()
                print("Send failure notification")
                send_notification_failure()
                #engine.execute(text('''EXEC [dbo].[ImportUpdateLog] @Param1 = ?, @Status = ?''', (curruser, 'Fail')).execution_options(autocommit=True))
        mail.Move(ProcessedFolder)
