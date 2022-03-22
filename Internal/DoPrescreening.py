# -*- coding: utf-8 -*-
"""
Created on Sat Mar  5 06:52:37 2022

@author: haasbroekj
"""

import webbrowser
import selenium
from selenium import webdriver
import time
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

s=Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=s)
driver.maximize_window()
driver.get("http://diyservices.nwu.ac.za/covid-19-pre-screening-service")
print("Modules imported")

studnumM = '34492275'
studnumW =  '37315145'
MPassW = 'hondZeus1!'
WPassW =  'HaasAw@1802'
print("variables declared")


def WaitRender(path):
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, path)))

def Checkbox(path):
    element = driver.find_element_by_xpath(path)
    driver.execute_script("arguments[0].scrollIntoView();",element)
    element.click()

def GetToken(number, password) :
    driver.maximize_window()
    driver.find_element_by_xpath('//*[@id="username"]').clear()
    driver.find_element_by_xpath('//*[@id="username"]').send_keys(number)
    driver.find_element_by_xpath('//*[@id="password"]').clear()
    driver.find_element_by_xpath('//*[@id="password"]').send_keys(password)
    driver.find_element_by_xpath('//*[@id="fm1"]/section[4]/input[4]').click()
    driver.switch_to.frame("myFrame")
    WaitRender('//*[@id="appName"]')
    Checkbox('//*[@id="disclaimerStartLayout"]/div[5]/span')
    WaitRender('//*[@id="fewer"]/span[2]/label')
    Checkbox('//*[@id="fewer"]/span[2]/label')
    Checkbox('//*[@id="cough"]/span[2]/label')
    Checkbox('//*[@id="throat"]/span[2]/label')
    Checkbox('//*[@id="breathing"]/span[2]/label')
    Checkbox('//*[@id="taste"]/span[2]/label')
    Checkbox('//*[@id="smell"]/span[2]/label')
    Checkbox('//*[@id="headache"]/span[2]/label')
    Checkbox('//*[@id="fatigue"]/span[2]/label')
    Checkbox('//*[@id="contact"]/span[2]/label')
    Checkbox('//*[@id="country"]/span[2]/label')
    Checkbox('//*[@id="positive"]/span[2]/label')
    Checkbox('//*[@id="submit"]')
    # time.sleep(3)
    # Checkbox('//*[@id="gwt-uid-47"]')

print("Functions created")

#webbrowser.open('http://diyservices.nwu.ac.za/covid-19-pre-screening-service')


WaitRender('//*[@id="username"]')

#GetToken(studnumM, MPassW)
GetToken(studnumW, WPassW)

