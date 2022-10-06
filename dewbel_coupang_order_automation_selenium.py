import time
import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import openpyxl
from openpyxl import Workbook
from openpyxl.comments import Comment
import html5lib
import pandas as pd
import numpy as np
import shutil

driver=webdriver.Chrome('chromedriver.exe')
#아이디 패스워드 :
ID='pm2100'
Password='cleanfeel1!'
#로그인 페이지
driver.get('https://xauth.coupang.com/auth/realms/seller/protocol/openid-connect/auth?response_type=code&client_id=supplier-hub&redirect_uri=https%3A%2F%2Fsupplier.coupang.com%2Flogin?returnUrl%3Dhttps%3A%2F%2Fsupplier.coupang.com%2Fscm%2Fpurchase%2Forder%2Flist&state=f7465c5f-da66-4d28-bfd2-57d02710d11e&login=true&scope=openid')
def xpath_element(xpath):
    return driver.find_element(By.XPATH,xpath)

def find_all_by_tagname(tagname, text):
    for element in driver.find_elements(By.TAG_NAME,tagname):
        print(element.text)
        if element.text==text:
            return element

xpath_element('//*[@id="root"]/main/form[1]/label[1]/input').send_keys(ID)
xpath_element('//*[@id="root"]/main/form[1]/label[2]/input').send_keys(Password)
xpath_element('//*[@id="root"]/main/form[1]/button').click()
'''//*[@id="rsNav030bab5f"]/li[3]/a' \
//*[@id="rsNavaaa02e0a"]/li[3]/a
#rsNavaaa02e0a > li:nth-child(3) > a

'''
try:
    xpath_element('//*[@id="app"]/div/div/a[2]').click()
except Exception:
    pass
driver.implicitly_wait(3)
driver.find_element(By.TAG_NAME,'button').click()


while True:

    try:
        for i in driver.find_elements(By.TAG_NAME,'button'):
            if i.text=="한국어":
                i.click();
                driver.implicitly_wait(2)
                break
        break
    except:
        pass
driver.implicitly_wait(3)
time.sleep(3)

for i in driver.find_elements(By.TAG_NAME, 'a'):
    #print(i.text)
    if i.text=="물류":i.click();break
for i in driver.find_elements(By.TAG_NAME, 'a'):
    if i.text=="발주리스트":
        i.click()
        break

xpath_element('//*[@id="purchaseOrderStatus"]/option[2]').click()
xpath_element('//*[@id="search"]').click()

elems=driver.find_elements(By.TAG_NAME,'table')
pd.read_html(elems[1].get_attribute('outerHTML'))[0]

elems=driver.find_elements(By.TAG_NAME,'table')
for i in elems[1].find_elements(By.TAG_NAME,'a'):
    try:
        int((i.text))
        i.send_keys(Keys.CONTROL+"\n")
    except : pass

driver.switch_to.window(driver.window_handles[1])

T.columns=T.columns.get_level_values(2)
T.iloc[::2,:].reset_index(drop=True)
