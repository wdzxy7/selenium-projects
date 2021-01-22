import re
import time
from selenium import webdriver
from bs4 import BeautifulSoup
from openpyxl import Workbook


def deal(s):
    pattern = re.compile(r'<[^>]*>')
    remove = re.findall(pattern, s)
    for i in remove:
        s = s.replace(i, '')
    return s


def get_max_page(html):
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find('div', {'class': 'pager-number'})
    num_lis = []
    for a in div.children:
        num = a.string
        try:
            num = int(num)
            num_lis.append(num)
        except:
            continue
    return max(num_lis)


if __name__ == '__main__':
    wb = Workbook()
    excel = wb.active
    excel['A1'] = 'opinions'
    count = 2
    url_list = []
    driver = webdriver.Chrome()
    urs = ['https://community.hihonor.com/unitedkingdom/forumplate/honor-magicbook/?forumId=393',
           'https://community.hihonor.com/france/forumplate/honor-magicbook/?forumId=379',
           'https://community.hihonor.com/germany/forumplate/honor-magicbook/?forumId=365',
           'https://community.hihonor.com/italy/forumplate/honor-magicbook/?forumId=397',
           'https://community.hihonor.com/spain/forumplate/honor-magicbook/?forumId=393']
    for ur in urs:
        driver.get(ur)
        time.sleep(5)
        html = driver.page_source
        max_page = get_max_page(html)
        signal = 0
        for i in range(2, 9):
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            div = soup.find_all('div', {'class': 'plate-posts-list'})
            soup = BeautifulSoup(str(div), 'html.parser')
            divs = soup.find_all('div', {'class': 'plate-comment-pre'})
            for div in divs:
                ID = div.get('id')
                name = deal(str(div)).replace('\n', '').strip().replace(' ', '-').replace('+', '').replace('[', '').replace(']', '')
                print(ID, name)
                url = 'https://community.hihonor.com/unitedkingdom/topicdetail/' + name + '/topicId-' + str(ID) + '/'
                print(url)
                if url in url_list:
                    signal = 1
                    break
                url_list.append(url)
            if signal == 1:
                break
            xpath = '//*[@id="plate-pager"]/div[1]/button[text()=\'>\']'
            button = driver.find_element_by_xpath(xpath)
            time.sleep(2)
            driver.execute_script("arguments[0].click();", button)
            time.sleep(5)
        for url in url_list:
            driver.get(url)
            time.sleep(3)
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            divs = soup.find_all('div', {'class': 'comments_content'})
            for div in divs:
                mess = deal(str(div)).replace('\n', '').replace('[', '').replace(']', '')
                if len(mess) == 0:
                    continue
                excel['A' + str(count)] = mess
                count = count + 1
                print(mess)
            divs = soup.find_all('div', {'class': 'comments_reply_tip'})
            for div in divs:
                mess = deal(str(div)).replace('\n', '').replace('[', '').replace(']', '')
                if len(mess) == 0:
                    continue
                excel['A' + str(count)] = mess
                count = count + 1
                print(mess)
    driver.close()
    wb.save('opinions.xlsx')