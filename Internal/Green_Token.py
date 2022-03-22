from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def WaitRender(path):
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, path)))
    
def Textbox(Text, path):
    element = driver.find_element_by_xpath(path)
    driver.execute_script("arguments[0].scrollIntoView();",element)
    element.send_keys(Text)
    
def Checkbox(path):
    element = driver.find_element_by_xpath(path)
    driver.execute_script("arguments[0].scrollIntoView();",element)
    element.click()
 
# file = open("Info","r")
# for x in range(1, 3):
#     line = file.readline()
#     txt = line.strip()
#     if x == 1:
#         username = txt
#     if x == 2:
#         password = txt

username = '27013588'
password = 'Styger21!'
driver = webdriver.Chrome()
driver.get("http://diyservices.nwu.ac.za/covid-19-pre-screening-service")
driver.maximize_window()
driver.switch_to.frame("myFrame")

# XPaths
User_path = '//*[@id="username"]'
Pass_path = '//*[@id="password"]'
Log_path = '//*[@id="fm1"]/section[4]/input[4]'
Render = '//*[@id="appName"]'
CB1_path = '//*[@id="disclaimerStartLayout"]/div[5]/span'
CB2_path = '//*[@id="vaccine"]/span[2]/label' 
CB3_path = '//*[@id="fewer"]/span[2]/label'
CB4_path = '//*[@id="cough"]/span[2]/label'
CB5_path = '//*[@id="throat"]/span[2]/label'
CB6_path = '//*[@id="breathing"]/span[2]/label'
CB7_path = '//*[@id="taste"]/span[2]/label'
CB8_path = '//*[@id="smell"]/span[2]/label'
CB9_path = '//*[@id="headache"]/span[2]/label'
CB10_path = '//*[@id="fatigue"]/span[2]/label'
CB11_path = '//*[@id="contact"]/span[2]/label'
CB12_path = '//*[@id="country"]/span[2]/label'
CB13_path = '//*[@id="positive"]/span[2]/label'
Final_path = '//*[@id="submit"]'

Textbox(username, User_path)
Textbox(password, Pass_path)
Checkbox(Log_path)
WaitRender(Render)
Checkbox(CB1_path)
WaitRender(CB2_path)

Checkbox(CB2_path)
Checkbox(CB3_path)
Checkbox(CB4_path)
Checkbox(CB5_path)
Checkbox(CB6_path)
Checkbox(CB7_path)
Checkbox(CB8_path)
Checkbox(CB9_path)
Checkbox(CB10_path)
Checkbox(CB11_path)
Checkbox(CB12_path)
Checkbox(CB13_path)
Checkbox(Final_path)

time.sleep(8)
driver.quit()
