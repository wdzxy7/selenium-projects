from selenium import webdriver
from bs4 import BeautifulSoup
import time
import requests
from requests.exceptions import RequestException
import re
from openpyxl import Workbook

wb = Workbook()
excel = wb.active
excel['A1'] = 'opinions'
count = 2


def get_html(url):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.162 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=20)
        response.encoding = 'utf-8'
        if response.status_code == 200:
            return response.text
        return None
    except RequestException:
        return None


def select_phone(html):
    global phone
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find('div', {'class': 'st-text'})
    soup = BeautifulSoup(str(div), 'html.parser')
    tds = soup.find_all('td')
    for td in tds:
        string = str(td)
        if phone in string.lower():
            soup = BeautifulSoup(str(string), 'html.parser')
            a = soup.find('a')
            back = a.get('href')
            print(a.get('href'))
            break
    return back


def get_opinions(html):
    global model
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find('div', {'class': 'makers'})
    soup = BeautifulSoup(str(div), 'html.parser')
    lis = soup.find_all('li')
    for li in lis:
        string = str(li).replace(' ', '')
        soup = BeautifulSoup(str(string), 'html.parser')
        span = soup.find('span')
        phone_name = span.contents[0]
        if str(phone_name).lower().replace(' ', '') == model.lower().replace(' ', ''):
            soup = BeautifulSoup(str(li), 'html.parser')
            a = soup.find('a')
            back = a.get('href')
            print(a.get('href'))
            break
    url = 'https://www.gsmarena.com/' + str(back).replace('-', '-reviews-')
    return url


def get_max_page(html):
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find('div', {'id': 'user-pages'})
    num_lis = []
    for a in div.children:
        num = a.string
        try:
            num = int(num)
            num_lis.append(num)
        except:
            continue
    return max(num_lis)


def crawl(html):
    global excel, count
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find('div', {'id': 'all-opinions'})
    soup = BeautifulSoup(str(div), 'html.parser')
    ps = soup.find_all('p', {'class', 'uopin'})
    for p in ps:
        soup = BeautifulSoup(str(p), 'html.parser')
        s = deal(str(p))
        opinion = [s.extract() for s in soup('span')]
        string = ''
        for i in opinion:
            string = string + str(i)
        opinion = deal(string)
        s = s.replace(opinion, '')
        print('result:')
        print(s)
        excel['A' + str(count)] = s
        count = count + 1
        print('--------------------------------------------')


def deal(s):
    pattern = re.compile(r'<[^>]*>')
    remove = re.findall(pattern, s)
    for i in remove:
        s = s.replace(i, '')
    return s


if __name__ == '__main__':
    print('输入手机品牌:')
    phone = input().lower()
    print('输入手机型号:')
    model = input()
    url = 'https://www.gsmarena.com/makers.php3'
    driver = webdriver.Chrome()
    driver.get(url)
    html = driver.page_source
    # driver.quit()
    # html = get_html(url)
    back = select_phone(html)
    url = 'https://www.gsmarena.com/' + back
    driver.get(url)
    html = driver.page_source
    url = get_opinions(html)
    driver.get(url)
    html = driver.page_source
    # html = get_html(url)
    max_page = get_max_page(html)
    print(max_page)
    driver.quit()
    for i in range(1, max_page + 1):
        page = 'p' + str(i) + '.php'
        this_url = url.replace('.php', page)
        print(this_url)
        try:
            html = get_html(this_url)
            crawl(html)
        except:
            continue
    save_name = phone + '_' + model
    wb.save(save_name + 'opinions.xlsx')