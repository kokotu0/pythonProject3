import os
from urllib.parse import urlparse
import urllib
import datetime
import pathlib
def selenium_order_list_save(coupang_supply_id,coupang_supply_password,start_date=(datetime.datetime.now()-datetime.timedelta(21)).strftime('%Y-%m-%d'),
                             path='C:\\Users\\Administrator\\Desktop\\한태희 파일\\test',):
    from selenium import webdriver

    from selenium.webdriver.common.by import By
    import pandas as pd
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import os
    try:
        os.makedirs(path)
        os.makedirs(f'{path}\\발주서')
    except:pass
    download_folder=path
    from selenium.webdriver.chrome.options import Options
    op=Options()
    op.add_experimental_option('prefs',{'download.default_directory':download_folder})
    driver=webdriver.Chrome(options=op)
    #아이디 패스워드 :


    #로그인 페이지
    url='https://xauth.coupang.com/auth/realms/seller/protocol/openid-connect/auth?response_type=code&client_id=supplier-hub&redirect_uri=https%3A%2F%2Fsupplier.coupang.com%2Flogin?returnUrl%3Dhttps%3A%2F%2Fsupplier.coupang.com%2Fscm%2Fpurchase%2Forder%2Flist&state=f7465c5f-da66-4d28-bfd2-57d02710d11e&login=true&scope=openid'

    driver.get(url)
    def xpath_element(xpath):
        return driver.find_element(By.XPATH,xpath)

    xpath_element('//*[@id="root"]/main/form[1]/label[1]/input').send_keys(coupang_supply_id)
    xpath_element('//*[@id="root"]/main/form[1]/label[2]/input').send_keys(coupang_supply_password)
    xpath_element('//*[@id="root"]/main/form[1]/button').click()

    driver.get('https://supplier.coupang.com/scm/purchase/order/list')#발주리스트


    table=pd.DataFrame()

    datetime.datetime.strftime(datetime.datetime.now(),'yyyy-mm-dd')
    end_date=datetime.datetime.now().strftime(format='%Y-%m-%d')
    end_date
    page=1
    while True:

        url=f'https://supplier.coupang.com/scm/purchase/order/list?page={page}&searchDateType=PURCHASE_ORDER_DATE&searchStartDate={start_date}&searchEndDate={end_date}'
        driver.get(url)

        parse=urlparse(driver.current_url)
        query=parse[4]
        urllib.parse.parse_qs(query)
        driver.execute_script("window.scrollBy(0,-50000)")
        '''
        발주서 전체 클릭 및 파일 다운로드
        
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[2]/div[2]/div[4]/table/thead/tr[1]/th[1]/div/label'))).click()
        driver.execute_script("$('#btn-download-po').click()")'''

        try:
            alert = driver.switch_to.alert
            alert.accept()
        except:
            pass
        # table=pd.concat([table,pd.read_html(driver.page_source)[1]])

        len(driver.find_elements(By.TAG_NAME,'table'))
        Ele=driver.find_element(By.XPATH,'//table[@class="scmTable basic marginT"]')
        new_table=pd.read_html(Ele.get_attribute('outerHTML'))

        if set(new_table[0].loc[0].values)=={'검색 결과가 없습니다.'}:break
        table=pd.concat([table,new_table[0]],ignore_index=True)
        page+=1

    print(table)

    table.to_excel(f'{path}\\PO_table.xlsx')
    '''
    다운받은 발주서를 압축풀기. 및 분류
    
    import shutil
    import zipfile
    import re
    for zip_file in [i for i in os.listdir(path) if '.zip' in i]:
        shutil.unpack_archive(f'{path}\\{zip_file}',extract_dir=f'{path}\\발주서',format='zip')
    for file in os.listdir(f'{path}\\발주서'):
        os.rename(f'{path}\\발주서\\{file}',f'{path}\\발주서\\{re.search(r'\d+',file).group(0)}.xlsx')
    '''
    # driver.find_element(By.XPATH,'//*[@id="pagination"]/ul/li[3]/a').get_attribute('href')
    driver.quit()
    del driver
    table.loc[:,[('발주','번호')]]=table[('발주','번호')].apply(lambda x  : re.search(r'\d+',x)[0])

    return table
import pandas as pd
import re
def table_interpret(PO_number,PO_table):
    PO_number='76255649'
    PO_table.loc[lambda x : x[('발주','번호')]==PO_number,('발주','운송')]




    pass

