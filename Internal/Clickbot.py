from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
driver = webdriver.Chrome(ChromeDriverManager().install())

Username = 'Janus'
Surname = 'Haasbroek'
email = 'janushaasbroek@gmail.com'
cellnum = '0769246784'

def WaitRender(path):
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, path)))

driver.get("https://uweb.unite180.com/")
print('OPENED BROWSER')

WaitRender('/html/body/app-root/div/mat-sidenav-container/mat-sidenav-content/div/div[1]/ng-component/div/mat-card/mat-card-content/div/button[1]')
driver.find_element_by_xpath('/html/body/app-root/div/mat-sidenav-container/mat-sidenav-content/div/div[1]/ng-component/div/mat-card/mat-card-content/div/button[1]').click()
selectElement = driver.find_element_by_xpath('//*[@id="mat-option-10"]/span')

#driver.navigate().refresh()

'''driver.find_element_by_xpath('//*[@id="wpcf7-f3436-p734-o2"]/form/p[1]/label/span/input').click()
print('clicked')
driver.find_element_by_xpath('//*[@id="wpcf7-f3436-p734-o2"]/form/p[1]/label/span/input').send_keys('Janus')
print('typed')
driver.find_element_by_xpath('//*[@id="wpcf7-f3436-p734-o2"]/form/p[6]/button').click()'''