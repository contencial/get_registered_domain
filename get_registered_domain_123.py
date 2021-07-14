import os
import re
import datetime
import gspread
from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from fake_useragent import UserAgent
from oauth2client.service_account import ServiceAccountCredentials

# Logger setting
from logging import getLogger, FileHandler, DEBUG
logger = getLogger(__name__)
today = datetime.datetime.now()
os.makedirs('./log', exist_ok=True)
handler = FileHandler(f'log/{today.strftime("%Y-%m-%d")}_result.log', mode='a')
handler.setLevel(DEBUG)
logger.setLevel(DEBUG)
logger.addHandler(handler)
logger.propagate = False

### functions ###
def write_domain_info(domain_info):
    SPREADSHEET_ID = os.environ['UNDER_CONTRACT_DOMAIN_SSID']
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheet.json', scope)
    gc = gspread.authorize(credentials)
    sheet = gc.open_by_key(SPREADSHEET_ID).worksheet('登録中ドメイン（123サーバー）')

    sheet.clear()
    cell_list = sheet.range(1, 1, len(domain_info) + 1, 3)
    i = 0
    for cell in cell_list:
        if (i == 0):
            cell.value = 'サーバー番号'
        elif (i == 1):
            cell.value = 'ドメイン名'
        elif (i == 2):
            cell.value = '契約中ドメイン\n存在確認'
        elif (i % 3 == 0):
            cell.value = domain_info[int(i / 3) - 1][0]
        elif (i % 3 == 1):
            cell.value = domain_info[int(i / 3) - 1][1]
        elif (i % 3 == 2):
            cell.value = domain_info[int(i / 3) - 1][2]
        i += 1
    sheet.update_cells(cell_list, value_input_option='USER_ENTERED')

    cell_list = sheet.range('D1:K1')
    cell_list[0].value = 'Size'
    cell_list[1].value = len(domain_info)
    cell_list[2].value = 'TRUE'
    cell_list[3].value = '=COUNTIF($C:$C, TRUE)'
    cell_list[4].value = 'FALSE'
    cell_list[5].value = '=COUNTIF($C:$C, FALSE)'
    cell_list[6].value = datetime.datetime.now().strftime('%Y-%m-%d')
    cell_list[7].value = '=HYPERLINK("https://member.123server.jp/servers/", "Go to 123Server")'
    sheet.update_cells(cell_list, value_input_option='USER_ENTERED')

def parse_contents(contents):
    tbody = contents.find_all("tbody")
    if (len(tbody) < 2):
        return None
    server_no = tbody[0].find("td").get_text()
    domain_list = tbody[1].find_all("a")
    for element in domain_list:
        domain_name = element.get_text()
        if re.search("m005b400", domain_name):
            continue
        yield [server_no, domain_name, f'=IF(COUNTIF(\'契約中ドメイン一覧\'!B:B, "{domain_name}"), TRUE, FALSE)']

def button_click(driver, button_text):
    buttons = driver.find_elements_by_tag_name("button")

    for button in buttons:
        if button.text == button_text:
            button.click()
            break

def get_domain_info():
    url = "https://member.123server.jp/members/login/"
    login = os.environ['SERVER123_USER']
    password = os.environ['SERVER123_PASS']
    webdriverPath = os.environ['WEBDRIVER_PATH']
    
    ua = UserAgent()
    logger.debug(f'123_server: UserAgent: {ua.chrome}')

    options = Options()
    options.add_argument('--headless')
    options.add_argument(ua.chrome)
    
    try:
        driver = webdriver.Chrome(executable_path=webdriverPath, options=options)
        
        driver.get(url)

        driver.find_element_by_id("MemberContractId").send_keys(login)
        driver.find_element_by_id("MemberPassword").send_keys(password)
        button_click(driver, "ログイン")
        
        logger.debug('123_server: login')
        sleep(3)
        
        driver.find_element_by_xpath('//a[@href="/servers/"]').click()
        
        logger.debug('123_server: go to server_list')
        sleep(3)

        paging = driver.find_element_by_xpath('//ul[@class="pagination"]').find_elements_by_tag_name("a")
        
        registered_domain_list = list()
        for i in range(len(paging)):
            if i < 1 or i > 3:
                continue
            logger.debug(f'123_server: page: {i}')
            driver.find_element_by_link_text(str(i)).click()
            sleep(10)
            for index in range(100):
                domain_list_button = driver.find_elements_by_link_text("ドメイン一覧")
                domain_list_button[index].click()
                server_no = 100 * (i - 1) + index + 1
                sleep(5)
                contents = BeautifulSoup(driver.page_source, "lxml")
                domain_chunk = list(parse_contents(contents))
                logger.debug(f'123_server: No {server_no}: {len(domain_chunk)}')
                registered_domain_list.extend(domain_chunk)
                driver.find_element_by_xpath('//a[@href="/servers/"]').click()
                sleep(3)
                driver.find_element_by_link_text(str(i)).click()
                sleep(3)

        logger.debug(f'123_server: total_list_number: {len(registered_domain_list)}')

        driver.close()
        driver.quit()

        return registered_domain_list
    except Exception as err:
        logger.debug(f'Error: 123_server: get_domain_info: {err}')
        exit(1)

### main_script ###
if __name__ == '__main__':

    try:
        logger.debug("123_server: start get_domain_info")
        domain_info = get_domain_info()
        logger.debug("123_server: start write_domain_list")
        write_domain_info(domain_info)
        exit(0)
    except Exception as err:
        logger.debug(f'123_server: {err}')
        exit(1)
