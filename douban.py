from selenium import webdriver
from bs4 import BeautifulSoup
import time
import requests
from requests.exceptions import RequestException
from openpyxl import Workbook


wb = Workbook()
excle = wb.active
excle['A1'] = 'name'
excle['B1'] = 'director'
excle['C1'] = 'country'
excle['D1'] = 'date'
excle['E1'] = 'genre'
excle['F1'] = 'actor'
excle['G1'] = 'better'
excle['H1'] = 'runtime'
excle['I1'] = 'abstract'
excle['J1'] = 'score'
count = 2


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


def crawl(html):
    soup = BeautifulSoup(html, 'html.parser')
    span = soup.find('span', {'property': 'v:itemreviewed'})
    try:
        name = span.contents[0]
    except:
        name = ''
    a = soup.find('a', {'rel': 'v:directedBy'})
    try:
        director = a.contents[0]
    except:
        director = ''
    span = soup.find('span', {'property': 'v:genre'})
    try:
        genre = span.contents[0]
    except:
        genre = ''
    span = soup.find('span', {'property': 'v:initialReleaseDate'})
    try:
        date = span.contents[0]
    except:
        date = ''
    span = soup.find('span', {'property': 'v:runtime'})
    try:
        runtime = span.contents[0]
    except:
        runtime = ''
    span = soup.find('span', {'property': 'v:summary'})
    try:
        abstract = span.contents[0]
    except:
        abstract = ''
    div = soup.find('div', {'class': 'rating_betterthan'})
    try:
        better = div.contents[1].contents[0]
    except:
        better = ''
    div = soup.find('div', {'class': 'rating_self clearfix'})
    try:
        score = div.contents[1].contents[0]
    except:
        score = ''
    a = soup.find('a', {'rel': 'v:starring'})
    try:
        actor = a.contents[0]
    except:
        actor = ''
    html = BeautifulSoup(str(soup.find('div', {'class': 'subject clearfix'})), 'html.parser')
    div = html.find('div', {'id': 'info'})
    pdd = 0
    country = ''
    for child in div.children:
        mes = str(child)
        if pdd == 1:
            country = mes.replace(' ', '')
            break
        if '制片国家/地区' in mes:
            pdd = 1
    print((name, director, genre, date, runtime, abstract, actor, score, better, country))
    store(name, director, genre, date, runtime, abstract, actor, score, better, country)


def store(name, director, genre, date, runtime, abstract, actor, score, better, country):
    global count, excle
    excle['A' + str(count)] = name
    excle['B' + str(count)] = director
    excle['C' + str(count)] = country
    excle['D' + str(count)] = date
    excle['E' + str(count)] = genre
    excle['F' + str(count)] = actor
    excle['G' + str(count)] = better
    excle['H' + str(count)] = runtime
    excle['I' + str(count)] = abstract
    excle['J' + str(count)] = score
    count = count + 1


if __name__ == '__main__':
    driver = webdriver.Chrome()
    url = 'https://movie.douban.com/explore#!type=movie&tag=%E6%81%90%E6%80%96&sort=recommend&page_limit=20&page_start=0'
    driver.get(url)
    count = 0
    while count != 10:
        xpath = '//a[text()=\'加载更多\']'
        try:
            driver.find_element_by_xpath(xpath).click()
        except:
           time.sleep(1.5)
           try:
                driver.find_element_by_xpath(xpath).click()
           except:
               break
        count += 1
    count = 2
    html = driver.page_source
    driver.quit()
    soup = BeautifulSoup(html, 'html.parser')
    a = soup.find_all('a', {'class': 'item'})
    for t in a:
        url = t.get('href')
        html = get_html(url)
        crawl(html)
    wb.save("douban_terror.xlsx")