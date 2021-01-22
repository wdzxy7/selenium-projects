import time
from selenium import webdriver
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook


wb = Workbook()
excel = wb.active
excel['A1'] = '视频链接'
excel['B1'] = '发布日期'
excel['C1'] = '观看量'
excel['D1'] = '评论数'
excel['E1'] = '点赞数（K）'
excel['F1'] = '点踩数'
count = 2


def move(driver):
    for i in range(9):
        js = "var q=document.documentElement.scrollTop=100000"
        driver.execute_script(js)
        time.sleep(1)


def get_url(a):
    return_list = []
    for i in a:
        back = i.get('href')
        ur = 'https://www.youtube.com' + back
        return_list.append(ur)
    return return_list


def crawl(html, url):
    global excel, count
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find('div', {'id': 'info', 'class': 'style-scope ytd-video-primary-info-renderer'})
    t_soup = BeautifulSoup(str(div), 'html.parser')
    span = t_soup.find('span', {'class': 'view-count style-scope yt-view-count-renderer'})
    view = span.contents[0]
    span = t_soup.find('yt-formatted-string', {'class': 'style-scope ytd-video-primary-info-renderer'})
    release_time = span.contents[0]
    yts = t_soup.find_all('yt-formatted-string', {'class': 'style-scope ytd-toggle-button-renderer style-text', 'id': 'text'})
    like = yts[0].contents[0]
    dislike = yts[1].contents[0]
    h2 = soup.find('h2', {'id': 'count', 'class': 'style-scope ytd-comments-header-renderer'})
    opinion = deal(str(h2))
    print(view, release_time, like, dislike, opinion)
    excel['A' + str(count)] = url
    excel['B' + str(count)] = release_time
    excel['C' + str(count)] = view
    excel['D' + str(count)] = opinion
    excel['E' + str(count)] = like
    excel['F' + str(count)] = dislike
    count += 1
    return release_time



def deal(s):
    pattern = re.compile(r'<[^>]*>')
    remove = re.findall(pattern, s)
    for i in remove:
        s = s.replace(i, '')
    return s


if __name__ == '__main__':
    url = 'https://www.youtube.com/c/GarenaFreeFireIndonesia/videos?view=0&flow=grid'
    driver = webdriver.Chrome()
    driver.get(url)
    move(driver)
    xpath = '//*[@id="contents"]'
    element = driver.find_element_by_xpath(xpath)
    html = element.get_attribute('innerHTML')
    driver.quit()
    soup = BeautifulSoup(html, 'html.parser')
    a = soup.find_all('a', {'class': 'yt-simple-endpoint style-scope ytd-grid-video-renderer'})
    url_list = get_url(a)
    t = 1
    driver = webdriver.Chrome()
    for url in url_list:
        print(len(url_list), t)
        print(url)
        driver.get(url)
        js = "var q=document.documentElement.scrollTop=800"
        driver.execute_script(js)
        time.sleep(2)
        html = driver.page_source
        release_time = crawl(html, url)
        if str(release_time).startswith('2019'):
            break
        time.sleep(0.4)
        t = t + 1
    wb.save('youtube.xlsx')