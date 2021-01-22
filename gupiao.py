import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from requests.exceptions import RequestException
import re
import time

last_time = '07-10'
sign1 = False
sign2 = False
title_list = []
time_list = []


def get_html(url):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.162 Safari/537.36'
        }
        response = requests.get(url, headers=headers)
        response.encoding = 'utf-8'
        if response.status_code == 200:
            return response.text
        return None
    except RequestException:
        return None


def crawl(body):
    global sign1, title_list, time_list
    soup = BeautifulSoup(str(body), 'html.parser')
    titles = soup.find_all('span', {'class': 'l3 a3'})
    for title in titles:
        title = deal(str(title))
        if title == '标题':
            continue
        title_list.append(title)
    times = soup.find_all('span', {'class': 'l5 a5'})
    for t in times:
        t = deal(str(t))
        if t == '发帖时间' or t == '最后更新':
            continue
        s = t.split(' ', 1)
        time_list.append(s[0])
        if s[0] == last_time:
            sign1 = True


def deal(s):
    pattern = re.compile(r'<[^>]*>')
    remove = re.findall(pattern, s)
    for i in remove:
        s = s.replace(i, '')
    return s


if __name__ == '__main__':
    wb = Workbook()
    excle = wb.active
    excle['A1'] = 'date'
    excle['B1'] = 'title sum'
    count = 2
    big_month = [1, 3, 5, 7, 8, 10, 12]
    small_month = [4, 6, 9, 11]
    store_title = []
    store_time = []
    dict1 = {}
    t1 = 7
    t2 = 13
    for i in range(1, 400):
        print(i)
        url = 'http://guba.eastmoney.com/list,zssh000001,f_' + str(i) + '.html'
        print(url)
        html = get_html(url)
        if html is None:
            continue
        soup = BeautifulSoup(html, 'html.parser')
        body = soup.find_all('div', {'id': 'articlelistnew'})
        crawl(body)
        time.sleep(0.1)
    t_time = time_list[0]
    titles = ''
    for title, t in zip(title_list, time_list):
        print(title, t)
        if dict1.__contains__(t):
            dict1[t] = dict1[t] + title
        else:
            dict1[t] = title
    for i in range(3):
        if t2 == 0:
            if t1 - 1 in big_month:
                t2 = 31
                t1 = t1 - 1
            elif t1 - 1 in small_month:
                t2 = 30
                t1 = t1 - 1
            else:
                t2 = 29
                t1 = t1 - 1
        if t1 == 0:
            t1 = 12
        if t1 < 10:
            if t2 < 10:
                date = '0' + str(t1) + '-0' + str(t2)
            else:
                date = '0' + str(t1) + '-' + str(t2)
        else:
            if t2 < 10:
                date = str(t1) + '-0' + str(t2)
            else:
                date = str(t1) + '-' + str(t2)
        try:
            title = dict1[date]
        except Exception as e:
            t2 -= 1
            continue
        excle['A' + str(count)] = date
        excle['B' + str(count)] = title
        del dict1[date]
        count = count + 1
        t2 -= 1
    wb.save("data crawler.xlsx")
