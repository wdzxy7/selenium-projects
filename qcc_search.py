import json
from selenium import webdriver
import time
from bs4 import BeautifulSoup
import xlrd
import xlutils.copy
import re

store_name_list = []
company_telephone_list = []
company_address_list = []
company_legal_person_list = []


def qcc_add_cookie(driver):
    try:
        xpath = '//*[@id="indexBonusModal"]/div/div/button'
        driver.find_element_by_xpath(xpath).click()
    except:
        pass
    with open('cookies(2).json', 'r', encoding='utf-8') as f:
        Cookies = json.loads(f.read())
    for cookie in Cookies:
        driver.add_cookie({
            "domain": cookie["domain"],
            "expirationDate": cookie['expirationDate'],
            "name": cookie['name'],
            "path": "/",
            "value": cookie['value'],
            "id": cookie['id']
        })
    driver.get('https://www.qcc.com/')


def get_telephon(html):
    soup = BeautifulSoup(html, 'html.parser')
    table1 = soup.find('table', {'class': 'ntable'})
    soup = BeautifulSoup(str(table1), 'html.parser')
    trs = soup.find_all('tr')
    tr = trs[1]
    soup = BeautifulSoup(str(tr), 'html.parser')
    tds = soup.find_all('td')
    money = tds[1]
    jy = deal(str(trs[8])).replace('经营范围', '')
    try:
        span = soup.find('span', {'style': 'color: #000;'})
        telephone = span.contents[0]
    except:
        telephone = ''
    try:
        a = soup.find('a', {'data-original-title': '查看地址'})
        address = a.contents[0]
    except:
        address = ''
    try:
        h2 = soup.find('h2', {'class': 'seo font-20'})
        legal_person = h2.contents[0]
    except:
        legal_person = ''
    tup = (telephone, address, legal_person, money, jy)
    return tup


def get_name(sheet):
    names = []
    message_sum = sheet.nrows + 1
    for i in range(1, message_sum):
        name = sheet.cell_value(i-1, 1)
        if name == '' or name == '名称':
            continue
        names.append(name)
    return names


def deal(s):
    pattern = re.compile(r'<[^>]*>')
    remove = re.findall(pattern, s)
    for i in remove:
        s = s.replace(i, '')
    return s


if __name__ == '__main__':
    # search_name = 'linyi paper'
    # path = 'ali_' + search_name + '.xls'
    path = 'testdata2.xlsx'
    data = xlrd.open_workbook(path)
    sheet = data.sheet_by_name("名单")
    # 复制表
    ws = xlutils.copy.copy(data)
    table = ws.get_sheet(0)
    # 获取搜索名称
    store_name_list = get_name(sheet)
    print(store_name_list)
    driver = webdriver.Chrome()
    # 登录企查查
    driver.get('https://www.qcc.com/')
    for i in range(1):
        time.sleep(1.5)
        try:
            qcc_add_cookie(driver)
            xpath = '//span[text()=\'登录 | 注册\']'
            try:
                driver.find_element_by_xpath(xpath)
                continue
            except:
                try:
                    xpath = '//*[@id="indexBonusModal"]/div/div/button'
                    driver.find_element_by_xpath(xpath).click()
                except:
                    pass
        except:
            pass
    search = driver.find_element_by_id('searchkey')
    search.clear()
    search.send_keys('沂南县华闰天元机械有限公司')
    element = driver.find_element_by_class_name('index-searchbtn')
    driver.execute_script("arguments[0].click();", element)
    countnum = 0
    del store_name_list[0]
    # 开始搜索
    for name in store_name_list:
        print(name)
        if name is None:
            company_telephone_list.append('None')
            company_address_list.append('None')
            company_legal_person_list.append('None')
            continue
        driver.find_element_by_id('headerKey')
        element = driver.find_element_by_xpath('//*[@id="headerKey"]')
        element.clear()
        element.send_keys(name)
        #search.clear()
        # search.send_keys(name)
        time.sleep(1.5)
        driver.find_element_by_xpath('//button[text()=\'查一下\']').click()
        back = driver.find_element_by_class_name('ma_h1').click()
        windows = driver.window_handles
        driver.switch_to.window(windows[-1])
        html = driver.page_source
        time.sleep(1.8)
        driver.close()
        driver.switch_to.window(windows[0])
        result = get_telephon(html)
        company_telephone_list.append(result[0])
        company_address_list.append(result[1])
        company_legal_person_list.append(result[2])
        print(result)
        countnum += 1
        time.sleep(1.7)
        if countnum == 100:
            countnum = 0
            driver.quit()
            driver = webdriver.Chrome()
            driver.get('https://www.qcc.com/')
            for i in range(3):
                try:
                    qcc_add_cookie(driver)
                    xpath = '//span[text()=\'登录 | 注册\']'
                    driver.find_element_by_xpath(xpath)
                    break
                except:
                    pass
            search = driver.find_element_by_id('searchkey')
            search.clear()
            search.send_keys('沂南县华闰天元机械有限公司')
            driver.find_element_by_class_name('index-searchbtn').click()
            time.sleep(60)
    count = 1
    print(len(store_name_list),
          len(company_legal_person_list),
          len(company_telephone_list), len(company_address_list))
    # 保存
    for i in range(len(company_address_list)):
        legal_person = company_legal_person_list[i]
        telephone = company_telephone_list[i]
        address = company_address_list[i]
        table.write(count, 13, legal_person)
        table.write(count, 14, telephone)
        table.write(count, 15, address)
        count = count + 1
    # save_name = search_name + '.xls'
    # ws.save('E://桌面文件//' + save_name)
    driver.quit()