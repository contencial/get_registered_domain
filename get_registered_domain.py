import os
import re
import codecs
import datetime
from ftplib import FTP
import gspread # to manipulate spreadsheet
from oauth2client.service_account import ServiceAccountCredentials # to access Google API

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
def get_ftp_server_info():
    SPREADSHEET_ID = os.environ['SERVERLIST_SSID']
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheet.json', scope)
    gc = gspread.authorize(credentials)
    sheet = gc.open_by_key(SPREADSHEET_ID).worksheet('ServerList')

    cell_list = sheet.range('I2:I301')
    ftp_server_list = [cell.value for cell in cell_list]
    return ftp_server_list

def get_existing_domain_list(server_no, host, user, passwd):
    ftp = FTP(
            host=host,
            user=user,
            passwd=passwd
        )

    items = ftp.mlsd('./public_html')
    domain_list = list()
    repatter = re.compile('.*\.\w+$')

    for filename, opt in items:
        if opt['type'] != 'dir' or filename == '.well-known':
            continue
        result = repatter.match(filename)
        if result:
            domain_list.append([server_no, filename, f'=IF(COUNTIF(\'契約中ドメイン一覧\'!B:B, "{filename}"), TRUE, FALSE)'])
    return domain_list

def write_registered_domain_list(domain_info):
    SPREADSHEET_ID = os.environ['UNDER_CONTRACT_DOMAIN_SSID']
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('spreadsheet.json', scope)
    gc = gspread.authorize(credentials)
    sheet = gc.open_by_key(SPREADSHEET_ID).worksheet('登録中ドメイン（FTPサーバー）')

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

### main_script ###
if __name__ == '__main__':

    user = os.environ['SERVER123_FTPUSER']
    passwd = os.environ['SERVER123_FTPPASS']
    server_no = 0
    registered_domain_info = list()
    try:
        ftp_server_list = get_ftp_server_info()
        for host in ftp_server_list:
            server_no += 1
            logger.info(f'{server_no}, {host}, {user}, {passwd}')
            domain_chunk = get_existing_domain_list(server_no, host, user, passwd)
            logger.info(f'server_no {server_no}: {len(domain_chunk)}')
            registered_domain_info.extend(domain_chunk)
        logger.info(len(registered_domain_info))
        logger.info("Write to Spreadsheet")
        write_registered_domain_list(registered_domain_info)
        logger.info("Finish")
    except Exception as err:
        logger.error(f'Error: {err}')
        exit(1)

