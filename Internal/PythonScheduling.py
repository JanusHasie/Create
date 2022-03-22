# -*- coding: utf-8 -*-
"""
Created on Fri Mar  4 09:55:14 2022

@author: haasbroekj
"""

#Import Moduels
import schedule
import datetime
import time

#Declare Variables
StartOnDate = '04/03/2022' # - Standard South African Format (dd/mm/yyyy)
DeployTime = [(StartOnDate, '07:20:00')]

def RunJob() :
    global DeployTime
    date = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    runTime = DeployTime[0] + " " + DeployTime[1]
    if DeployTime and date == str(runTime) :
        #do the task
        print("Hello World")
        
schedule.every(10).minutes.do(RunJob())

while True :
    schedule.run_pending()
    time.sleep(1)

    
#s= sched.scheduler(time.time, time.sleep)
#def print_time(a= 'default') :
#    print("From print_time", time.time(), a)
    
#def print_some_times() :
#    print(time.time())
#    s.enter()