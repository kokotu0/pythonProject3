from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

driver=webdriver.Chrome('chromedriver')


driver.get('https://www.g2b.go.kr/pt/menu/selectSubFrame.do?framesrc=/pt/menu/frameTgong.do?url=https://www.g2b.go.kr:8101/ep/tbid/tbidFwd.do'
           '')
#검색과정
driver.implicitly_wait(1)
driver.switch_to.frame('sub')
driver.switch_to.frame('main')
A=driver.find_element(By.XPATH,'//*[@id="bidNm"]')
A.send_keys("전략")

date_entry=driver.find_element(By.XPATH,'//*[@id="fromBidDt"]')
date_entry.click()

date_entry.send_keys(Keys.CONTROL+"A")
date_entry.send_keys(Keys.DELETE)
date_entry.send_keys("2022/09/14")
driver.implicitly_wait(1)
date_entry=driver.find_element(By.XPATH,'//*[@id="toBidDt"]')
date_entry.click()
date_entry.send_keys(Keys.CONTROL+"A")
date_entry.send_keys(Keys.DELETE)
date_entry.send_keys("2022/09/14")

driver.find_element(By.XPATH,'//*[@id="taskClCds5"]').click()
driver.find_element(By.XPATH,'//*[@id="recordCountPerPage"]/option[5]').click()
driver.find_element(By.XPATH,'//*[@id="buttonwrap"]/div/a[1]').click()
#resultForm > div.results > table > tbody > tr:nth-child(1) > td:nth-child(2) > div > a
#resultForm > div.results > table > tbody > tr:nth-child(1) > td:nth-child(4) > div > a
#resultForm > div.results > table > tbody > tr:nth-child(2) > td:nth-child(4) > div > a
'''                                                행넘버       열넘버                    텍스트지정'''
driver.implicitly_wait(1)
driver.switch_to.frame(driver.find_element(By.TAG_NAME('sub')))
driver.switch_to.frame('main')
# print(driver.find_element((By.CSS_SELECTOR,'#resultForm > div.results > table > tbody > tr:nth-child(3) > td:nth-child(4) > div > a')))
# print(driver.find_element((By.CSS_SELECTOR,'#resultForm > div.results > table > tbody > tr:nth-child(3) > td:nth-child(4) > div > a')).text)
n=0
while True:
    n+=1
    try:
        print(driver.find_element(By.XPATH,'//*[@id="resultForm"]/div[2]/table/tbody/tr[{0}]/td[4]/div/a'.format(n)).text)
    except Exception as error:
        break