from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException
from pyquery import PyQuery as pq
import json
import pymongo
from openpyxl import Workbook

client = pymongo.MongoClient(host='localhost')
MONGO_DB = 'stock'
MONGO_COLLECTION = '沪市A股数据'
db = client[MONGO_DB]
lines = []
MAX_PAGE = 49

option = webdriver.ChromeOptions()
option.add_argument('--user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.108 Safari/537.36"')
browser = webdriver.Chrome(options=option)
wait = WebDriverWait(browser, 5)

def get_page(page):
    print('正在爬取第', page, '页')
    try:
        url = 'http://quote.stockstar.com/stock/sha_3_1_1.html'
        browser.get(url)
        if page > 1:
            input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.fenye .page_input')))
            submit = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.fenye a:last-child > em')))
            input.clear()
            input.send_keys(page)
            submit.click()
        wait.until(EC.text_to_be_present_in_element((By.CSS_SELECTOR, '.fenye span > em'), str(page)))
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.tbody_right tr')))
        get_data()
    except TimeoutException:
        get_page(page)


def get_data():
    html = browser.page_source
    doc = pq(html)
    items = doc('#datalist tr').items()
    for item in items:
        data = {}
        data['代码'] = item.find('td:nth-child(1)').text()
        data['简称'] = item.find('td:nth-child(2)').text()
        data['最新价'] = item.find('td:nth-child(3)').text()
        data['涨跌幅'] = item.find('td:nth-child(4)').text()
        data['涨跌额'] = item.find('td:nth-child(5)').text()
        data['五分钟涨幅'] = item.find('td:nth-child(6)').text()
        data['成交量'] = item.find('td:nth-child(7)').text()
        data['成交额'] = item.find('td:nth-child(8)').text()
        data['换手率'] = item.find('td:nth-child(9)').text()
        data['振幅'] = item.find('td:nth-child(10)').text()
        data['量比'] = item.find('td:nth-child(11)').text()
        data['市盈率'] = item.find('td:nth-child(12)').text()
#        print(data)
        datalist = [data['代码'], data['简称'], data['最新价'], data['涨跌幅'],
            data['涨跌额'], data['五分钟涨幅'], data['成交量'], data['成交额'],
            data['换手率'], data['振幅'], data['市盈率'], data['量比']]
        lines.append(datalist)
        save_data(lines)


def save_data(lines):
    #保存到Excel
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(['代码','简称', '最新价', '涨跌幅', '涨跌额','五分钟涨幅', '成交量', '成交额', '换手率', '振幅', '量比', '市盈率'])
    for line in lines:
        worksheet.append(line)
    workbook.save('沪市A股数据.xlsx')
#    保存为txt格式
#    with open('data.txt', 'a', encoding='utf-8') as f:
#        f.write(json.dumps(data, ensure_ascii=False) + '\n')
#    保存到mongoDB
#    try:
#        if db[MONGO_COLLECTION].insert(data):
#            print('存储到MONGODB成功')
#    except Exception:
#        print('存储到MONGODB失败')

def main():
    for i in range(1, MAX_PAGE + 1):
        get_page(i)

if __name__ == '__main__':
    main()
