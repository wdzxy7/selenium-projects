from bs4 import BeautifulSoup
from requests.exceptions import RequestException
import requests
import re
from openpyxl import Workbook
import time
from selenium import webdriver


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


def whether_active(name):
    pattern = re.compile(r'.互动')
    if re.match(pattern, name):
        return 'yes'
    else:
        return 'no'


def get_time(soup):
    tag_span = soup.find_all('span', {'class': 'time'})
    pattern = re.compile(r'</i>(.*?)</span>$')
    time_list = []
    for i in tag_span:
        time = re.findall(pattern, str(i).replace(' ', '').replace('\n', ''))
        if len(time) > 0:
            print(time)
            time_list.append(time)
    return time_list


def get_aid(url):
    html = get_html(url)
    soup = BeautifulSoup(html, "html.parser")
    meta = soup.find_all('meta', {'property': 'og:url'})
    pattern = re.compile(r'/av(.*?)/')
    aid = re.findall(pattern, str(meta))
    return aid


def get_like(aid):
    url = 'https://api.bilibili.com/x/web-interface/archive/stat?aid=' + str(aid[0])
    html = get_html(url)
    pattern = re.compile(r'\"view\":(.*?),\"danmaku\"')
    view = re.findall(pattern, html)
    pattern = re.compile(r'\"danmaku\":(.*?),"reply\"')
    dm = re.findall(pattern, html)
    pattern = re.compile(r'\"favorite\":(.*?),\"coin\"')
    favorite = re.findall(pattern, html)
    pattern = re.compile(r'\"coin\":(.*?),\"share\"')
    coin = re.findall(pattern, html)
    pattern = re.compile(r'\"share\":(.*?),\"like')
    share = re.findall(pattern, html)
    pattern = re.compile(r'\"like\":(.*?),\"now_rank')
    like = re.findall(pattern, html)
    return view, dm, like, coin, favorite, share


def initial_excle(excle):
    excle['A1'] = '视频标题'
    excle['B1'] = 'BV号'
    excle['C1'] = '发布时间'
    excle['D1'] = '是否是互动视频'
    excle['E1'] = '播放量'
    excle['F1'] = '弹幕量'
    excle['G1'] = '点赞'
    excle['H1'] = '投币'
    excle['I1'] = '收藏'
    excle['J1'] = '转发'


def write_exlce(excle, count, title, time, BV, active, view, dm, like, coin, favorite, share):
    key = 'A' + count
    excle[key] = title
    key = 'B' + count
    excle[key] = BV
    key = 'C' + count
    excle[key] = time
    key = 'D' + count
    excle[key] = active
    key = 'E' + count
    excle[key] = view
    key = 'F' + count
    excle[key] = dm
    key = 'G' + count
    excle[key] = like
    key = 'H' + count
    excle[key] = coin
    key = 'I' + count
    excle[key] = favorite
    key = 'J' + count
    excle[key] = share


def get_first_html(url):
    driver = webdriver.Chrome()
    driver.get(url)
    time.sleep(3)
    return driver.page_source


if __name__ == '__main__':
    wb = Workbook()
    excle = wb.active
    initial_excle(excle)
    up_num = '11616487'
    max_page = 3
    count = 1
    for page in range(1, max_page + 1):
        url = 'https://space.bilibili.com/' + str(up_num) + '/video?tid=0&page=' + str(page) + '&keyword=&order=pubdate'
        html = get_first_html(url)
        soup = BeautifulSoup(html, "html.parser")
        print(soup)
        div = soup.find_all('a', {'class': 'title', 'target': '_blank'})
        BV_list = set()
        title_list = set()
        active_list = []
        dm_list = []
        view_list = []
        coin_list = []
        like_list = []
        favorite_list = []
        share_list = []
        url_list = set()
        for i in div:
            time.sleep(0.5)
            last = str(i.get('href'))
            pattern = re.compile(r'.*BV(.*?)$')
            BV = re.findall(pattern, last)
            BV_list.add(str(BV))
            url = 'https:' + last
            url_list.add(url)
            title = str(i.get('title'))
            title_list.add(title)
            print(title)
            active = whether_active(title)
            print(active)
            active_list.append(active)
        for url in url_list:
            aid = get_aid(url)
            view, dm, like, coin, favorite, share = get_like(aid)
            dm_list.append(dm)
            view_list.append(view)
            coin_list.append(coin)
            like_list.append(like)
            favorite_list.append(favorite)
            share_list.append(share)
        time_list = get_time(soup)
        BV_list = list(BV_list)
        title_list = list(title_list)
        for i in range(len(BV_list)):
            count = count + 1
            title = str(title_list[i]).replace('[', '').replace(']', '').replace('\'', '')
            times = str(time_list[i]).replace('[', '').replace(']', '').replace('\'', '')
            BV = str(BV_list[i]).replace('[', '').replace(']', '').replace('\'', '')
            active = str(active_list[i]).replace('[', '').replace(']', '').replace('\'', '')
            view = str(view_list[i]).replace('[', '').replace(']', '').replace('\'', '')
            dm = str(dm_list[i]).replace('[', '').replace(']', '').replace('\'', '')
            like = str(like_list[i]).replace('[', '').replace(']', '').replace('\'', '')
            coin = str(coin_list[i]).replace('[', '').replace(']', '').replace('\'', '')
            favorite = str(favorite_list[i]).replace('[', '').replace(']', '').replace('\'', '')
            share = str(share_list[i]).replace('[', '').replace(']', '').replace('\'', '')
            print(title, times, BV, active, view, dm, like, coin, favorite, share)
            write_exlce(excle, str(count), title, times, BV, active, view, dm, like, coin, favorite, share)
    wb.save("BiliBili.xlsx")