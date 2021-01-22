import time
from selenium import webdriver
import selenium
from openpyxl import Workbook
import re
import xlrd
from bs4 import BeautifulSoup


def deal(s):
    pattern = re.compile(r'<[^>]*>')
    remove = re.findall(pattern, s)
    for i in remove:
        s = s.replace(i, '')
    return s.replace('\n', '')


def get_url():
    read_excel = xlrd.open_workbook('反馈.xlsx')
    sheet = read_excel.sheet_by_name("Sheet1")
    message_sum = sheet.nrows + 1
    for i in range(2, message_sum):
        url = sheet.cell_value(i - 1, 0)
        url_list.append(url)


if __name__ == '__main__':
    url_list = []
    get_url()
    wb = Workbook()
    excle = wb.active
    excle['A1'] = 'comments'
    excle['B1'] = 'time'
    c_count = 2
    t_count = 2
    driver = webdriver.Chrome()
    for url in url_list:
        if len(url) == 0:
            continue
        driver.get(url)
        time.sleep(10)
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        div = soup.find('div', {'class': 'list_box'})
        soup = BeautifulSoup(str(div), 'html.parser')
        comments = soup.find_all('div', {'class': 'WB_text'})
        for comment in comments:
            com = deal(str(comment))
            com = re.sub(r'(.*?)：', '', com)
            excle['A' + str(c_count)] = com
            c_count += 1
            print(com)
        times = soup.find_all('div', {'class': 'WB_from S_txt2'})
        for t in times:
            com_time = deal(str(t))
            excle['B' + str(t_count)] = com_time
            print(com_time)
            t_count += 1
    driver.quit()
    wb.save("sina_comment3.xlsx")