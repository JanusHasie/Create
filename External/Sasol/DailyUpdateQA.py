# Imports
import os
import win32com.client
from datetime import datetime,timedelta
import time, re, shutil, os, pandas as pd
from sqlalchemy import create_engine, text
import datetime
import pyodbc as po
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
SendMailTo = 'haasbroekj@fourier.co.za', 'voudtshoornm@fourier.co.za'
print('Mail Details Updated')

# SQL Connection
engine = create_engine("mssql+pymssql://{user}:{pw}@197.189.232.50/{db}"
                       .format(user="sa", pw="NewFAsys098!", db="Sasol_Transport_DB"))
conn = po.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=197.189.232.50; DATABASE=[Sasol_Transport_DB]; UID=sa; PWD=NewFAsys098!')
cursor = conn.cursor()
print('SQL Engine Created')

#Run Stored Proc function
def execute_procedure(session, proc_name, has_param, params):
    if has_param = 'y'
        sql_params = ",".join(["@{0}={1}".format(name, value) for name, value in params.items()])
        sql_string = """
        DECLARE @return_value int;
        EXEC @return_value = [dbo].[{proc_name}] {params};
        SELECT 'return_value' = @return_value;
        """.format(proc_name = proc_name, params= sql_params)
        return engine.execute(sql_string).fetchall()
    elif has_param = 'n'
        sql_string = """
        EXEC @return_value = [dbo].[{proc_name}];
        """.format(proc_name = proc_name)
        return engine.execute(sql_string).fetchall()
    
    

# Email Creation
def send_notification_success(email):
    outlook_mail = win32com.client.Dispatch('outlook.application')
    mail_ = outlook_mail.CreateItem(0)
    mail_.To = email
    mail_.Subject = 'Data Update - Pass'
    mail_.body = 'The database was successfully updated'
    mail_.Send()


def send_notification_failure(email):
    outlook_mail = win32com.client.Dispatch('outlook.application')
    mail_ = outlook_mail.CreateItem(0)
    mail_.To = email
    mail_.Subject = 'Data Update - Fail'
    mail_.body = 'The database failed to update'
    mail_.Send()


def open_outlook():
    try:
        subprocess.call(['C:\\Program Files\\Microsoft Office\\root\\Office16\\Outlook.exe'])
        os.system("C:\\Program Files\\Microsoft Office\\root\\Office16\\Outlook.exe")
        print('Outlook open')
    except:
        print("Outlook didn't open successfully")


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
            # Run the stored procedure
            engine.execute(text('''EXEC [dbo].[AnalyseFY22]''').execution_options(autocommit=True))
            #engine.execute(text('''EXEC [dbo].[AnalyseFY22]''')).execution_options(autocommit=True)
            print('Daily Update Completed')
            if (flag == 1):
                for mail in SendMailTo :
                    send_notification_success(mail)
            else:
                open_outlook()
                for mail in SendMailTo :
                    send_notification_failure(mail)               
        except Exception as e:
            print(str(e) + ' | File NOT saved [{}]'.format(temp_filename))
            if (flag == 1):
                for mail in SendMailTo :
                    send_notification_failure(mail)
            else:
                open_outlook()
                for mail in SendMailTo :
                    send_notification_failure(mail)
        mail.Move(ProcessedFolder)
