import json
from selenium import webdriver
import time
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
from requests.exceptions import RequestException
import requests
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait

main_country_list = []
store_url_list = []
store_name_list = []
year_list = []
english_name_list = []
chinese_name_list = []
main_product_list = []
sold_list = []
money_list = []
url_list = []
manager_name_list = []
department_list = []
job_list = []
telephone_list = []
mobile_list = []
fax_list = []
company_telephone_list = []
company_address_list = []
company_legal_person_list = []
stored_url = []


class Warn(Exception):
    pass


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


def log(driver):
    xpath = '//a[@data-val=\'ma_signin\']'
    element = driver.find_element_by_xpath(xpath)
    driver.execute_script("arguments[0].click();", element)
    # ActionChains(driver).move_to_element(element).perform()
    # element.click()
    driver.implicitly_wait(20)
    xpath = '//input[@name=\'loginId\']'
    driver.find_element_by_xpath(xpath).send_keys('632080907@qq.com')
    xpath = '//input[@name=\'password\']'
    driver.find_element_by_xpath(xpath).send_keys('Ln632080907')
    xpath = '//span[@id=\'nc_1_n1z\']'
    try:
        span = driver.find_element_by_xpath(xpath)
        action = ActionChains(driver)
        action.click_and_hold(span).perform()
        # tab.drag_and_drop_by_offset(span, 1000, 0).perform()
        for i in range(200):
            try:
                action.move_by_offset(200, 0).perform()
            except:
                break
            action.reset_actions()
            time.sleep(0.1)
    except:
        pass
    driver.find_element_by_id('fm-login-submit').click()


def crawl(html):
    global driver2, year_list, english_name_list, main_product_list, sold_list, money_list, url_list, chinese_name_list,\
        manager_name_list, department_list, job_list, telephone_list, mobile_list, fax_list

    soup = BeautifulSoup(html, 'html.parser')
    divs = soup.find_all('div', {'class': 'f-icon m-item'})
    for div in divs:
        country = []
        soup = BeautifulSoup(str(div), 'html.parser')
        div = soup.find('div', {'class': 's-gold-supplier-year-icon'})
        try:
            year = div.contents[0]
        except:
            year = 'None'
        h2 = soup.find('h2', {'class': 'title ellipsis'})
        url = re.findall(r'href=\"(.*?)\"', str(h2))[0]
        english_name = deal(str(h2)).replace('\n', '')
        try:
            div = soup.find('div', {'class': 'value ellipsis ph'})
            main_product = deal(str(div))
        except:
            main_product = 'None'
        try:
            div = soup.find('div', {'class': 'lab'})
            sold = deal(str(div)).replace(' ', '').replace('\n', '')
        except:
            sold = 'None'
        div = soup.find('div', {'class': 'num'})
        try:
            money = re.findall(r'</i>(.*?)</div>', str(div))[0]
        except:
            money = 'None'
        try:
            spans = soup.find_all('span', {'class': 'ellipsis search'})
            for span in spans:
                coun = span.contents[0]
                if '%' in str(coun):
                    country.append(coun)
            country = tuple(country)
        except:
            country = 'None'
        if re.match(r'^(Shandong|Jinan|Qingdao|Zibo|Zaozhuang|Dongying|Yantai|Weifang|Jining|Taian'
                    r'|Weihai|Rizhao|Binzhou|Dezhou|Liaocheng|Linyi|Heze)', english_name):
            chinese_name = crawl_name(url)
            english_name_list.append(english_name)
            chinese_name_list.append(chinese_name)
            year_list.append(year)
            main_product_list.append(main_product)
            sold_list.append(sold)
            money_list.append(money)
            url_list.append(url)
            main_country_list.append(str(country))
            #print(url)
            print(year, end='\n****************\n')
            print(english_name, end='\n***************\n')
            #print(chinese_name)
            #print(country, end='\n*****************\n')
            #print(main_product, end='\n*****************\n')
            print(sold, end='\n***************************\n')
            print(money, end='\n*************************\n')
            print('-----------------------------------------------------------------------------------')


def deal(s):
    pattern = re.compile(r'<[^>]*>')
    remove = re.findall(pattern, s)
    for i in remove:
        s = s.replace(i, '')
    return s


def crawl_name(url):
    try:
        html = get_html(url)
        soup = BeautifulSoup(html, 'html.parser')
    except:
        return None
    try:
        a = soup.find_all('a', {'data-spm': 'dassicon'})
        if len(a) == 0:
            raise Warn()
    except Warn:
        try:
            a = soup.find_all('a', {'data-spm': 'davicon'})
            if len(a) == 0:
                raise Warn()
        except Warn:
            try:
                a = soup.find_all('a', {'class': 'sesame-click-target'})
            except:
                return None
    if len(a) == 0:
        return None
    try:
        url = url.replace('company_profile.html#top-nav-bar', '') + a[0].get('href')
    except:
        return None
    html = get_html(url)
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find_all('div', {'class': 'ui2-balloon ui2-balloon-bl ui2-shadow-normal registration-name-tip'})
    div = str(div)
    name = re.findall(r'name "(.*?)" on', div)
    if len(name) == 0:
        return None
    return name[0]


def get_name(driver, url):
    xpath = '//span[@title=\'Contacts\']'
    try:
        driver.find_element_by_xpath(xpath).click()
    except:
        time.sleep(3)
        try:
            driver.find_element_by_xpath(xpath).click()
        except:
            tup = ('', '', '', '', '', '')
            return tup

    xpath = '//a[text()=\'View details\']'
    wait = WebDriverWait(driver, 4, 0.2)
    try:
        wait.until(lambda dribver: driver.find_element_by_xpath('//a[text()=\'View details\']'))
        for i in range(40):
            try:
                element = driver.find_element_by_xpath(xpath)
                driver.execute_script("arguments[0].click();", element)
                break
            except Exception as e:
                pass
    except:
        pass
    time.sleep(1.5)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    name_div = soup.find('div', {'class': 'contact-name'})
    try:
        name = name_div.contents[0]
    except:
        name = ''
    department_div = soup.find('div', {'class': 'contact-department'})
    try:
        department = department_div.contents[0]
    except:
        department = ''
    job_div = soup.find('div', {'class': 'contact-job'})
    try:
        job = job_div.contents[0]
    except:
        job = ''
    table = str(soup.find('table', {'class': 'info-table'}))
    soup = BeautifulSoup(table, 'html.parser')
    trs = soup.find_all('tr')
    telephone = ''
    mobile = ''
    fax = ''
    for tr in trs:
        try:
            str1 = str(tr.contents[0])
            str2 = str(tr.contents[1])
        except:
            continue
        if re.search(r'Telephone', str1):
            telephone = str2.replace('<td>', '').replace('</td>', '')
        elif re.search(r'Mobile Phone', str1):
            mobile = str2.replace('<td>', '').replace('</td>', '')
        elif re.search(r'Fax', str1):
            fax = str2.replace('<td>', '').replace('</td>', '')
        else:
            break
    telephone = re.sub(r'^(86-|0086-|0086)', '', telephone)
    mobile = re.sub(r'^(86-|0086-|0086)', '', mobile)
    fax = re.sub(r'^(86-|0086-|0086)', '', fax)
    tup = (name, department, job, telephone, mobile, fax)
    return tup


def qcc_add_cookie(driver):
    try:
        xpath = '//*[@id="indexBonusModal"]/div/div/button'
        driver.find_element_by_xpath(xpath).click()
    except:
        pass
    with open('cookies.json', 'r', encoding='utf-8') as f:
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
    tup = (telephone, address, legal_person)
    return tup


def creat_wb(excle):
    excle['A1'] = '公司名称'
    excle['B1'] = '链接'
    excle['C1'] = '主营产品'
    excle['D1'] = '年限'
    excle['E1'] = '主要市场'
    excle['F1'] = '账单和金额'
    excle['G1'] = '联系人'
    excle['H1'] = '联系人职位'
    excle['I1'] = '联系人部门'
    excle['J1'] = '电话'
    excle['K1'] = '手机'
    excle['L1'] = '传真'
    excle['M1'] = '公司法人代表'
    excle['N1'] = '公司电话'
    excle['O1'] = '公司地址'
    excle['P1'] = '英文名字'


def t_log(driver):
    u = 'https://passport.alibaba.com/icbu_login.htm?return_url=https%3A%2F%2Fi.alibaba.com%2Findex.htm%3Fspm%3Da2700.galleryofferlist.scGlobalHomeHeader.7.c402e202wt2y9u%26tracelog%3Dhd_ma'
    driver.get(u)
    driver.implicitly_wait(20)
    try:
        xpath = '//input[@name=\'loginId\']'
        driver.find_element_by_xpath(xpath).send_keys('632080907@qq.com')
        xpath = '//input[@name=\'password\']'
        driver.find_element_by_xpath(xpath).send_keys('Ln632080907')
        driver.find_element_by_id('fm-login-submit').click()
        time.sleep(3)
    except:
        pass


def creat_ali_wb(excle):
    excle['A1'] = '公司名称'
    excle['B1'] = '英文名字'
    excle['C1'] = '链接'
    excle['D1'] = '主营产品'
    excle['E1'] = '年限'
    excle['F1'] = '主要市场'
    excle['G1'] = '账单和金额'
    excle['H1'] = '联系人'
    excle['I1'] = '联系人职位'
    excle['J1'] = '联系人部门'
    excle['K1'] = '电话'
    excle['L1'] = '手机'
    excle['M1'] = '传真'
    excle['N1'] = '公司法人代表'
    excle['O1'] = '公司电话'
    excle['P1'] = '公司地址'


def deep_crawl_main():
    wb2 = Workbook()
    ali_excle = wb2.active
    creat_ali_wb(ali_excle)
    search_name = 'linyi paper'
    driver = webdriver.Chrome()
    driver.maximize_window()
    url = 'https://www.alibaba.com/trade/search?fsb=y&IndexArea=company_en&CatId=&SearchText=' + search_name
    driver.get(url)
    log(driver)
    while True:
        js = "var q=document.documentElement.scrollTop=100000"
        driver.execute_script(js)
        xpath = '//div[@class=\'grid grid-c3-s5e6\']'
        res = driver.find_element_by_xpath(xpath)
        html = res.get_attribute('innerHTML')
        if html is None:
            continue
        crawl(html)
        xpath = '//a[@class=\'next\']'
        try:
            next_button = driver.find_element_by_xpath(xpath)
        except:
            break
        next_button.click()
        time.sleep(1)
    driver.quit()
    print(len(store_name_list), len(store_url_list), len(chinese_name_list), len(main_product_list), len(year_list), len(main_country_list),
          len(money_list), len(manager_name_list)
          , len(job_list), len(department_list), len(telephone_list), len(mobile_list), len(fax_list),
          len(company_legal_person_list),
          len(company_telephone_list), len(company_address_list))
    driver = webdriver.Chrome()
    driver.get(url)
    log(driver)
    time.sleep(4)
    stored_name = set()
    search_sum = 0
    length = len(chinese_name_list)
    del_count = 0
    for i in range(length):
        if chinese_name_list[i] not in stored_name:
            company_name = chinese_name_list[i]
            print(chinese_name_list[i], url_list[i])
            stored_name.add(chinese_name_list[i])
            url = url_list[i]
            try:
                driver.get(url)
            except:
                result = ['', '', '', '', '', '']
                store_name_list.append(company_name)
                store_url_list.append(url)
                manager_name_list.append(result[0])
                department_list.append(result[1])
                job_list.append(result[2])
                telephone_list.append(result[3])
                mobile_list.append(result[4])
                fax_list.append(result[5])
                continue
            result = get_name(driver, url)
            print(result)
            store_name_list.append(company_name)
            store_url_list.append(url)
            manager_name_list.append(result[0])
            department_list.append(result[1])
            job_list.append(result[2])
            telephone_list.append(result[3])
            mobile_list.append(result[4])
            fax_list.append(result[5])
            time.sleep(0.5)
            # 查询多少次检测一下是否还处于登录状态
            if search_sum == 10:
                driver.quit()
                driver = webdriver.Chrome()
                search_url = 'https://www.alibaba.com/trade/search?fsb=y&IndexArea=company_en&CatId=&SearchText=' + search_name
                driver.get(search_url)
                try:
                    driver.implicitly_wait(10)
                    log(driver)
                    time.sleep(3)
                except:
                    pass
                search_sum = 0
            search_sum = search_sum + 1
        else:
            del english_name_list[i - del_count]
            del main_country_list[i - del_count]
            del main_product_list[i - del_count]
            del year_list[i - del_count]
            del money_list[i - del_count]
            del_count = del_count + 1
        print(len(store_name_list), len(store_url_list), len(main_product_list), len(year_list), len(main_country_list),
              len(money_list), len(manager_name_list)
              , len(job_list), len(department_list), len(telephone_list), len(mobile_list), len(fax_list),
              len(company_legal_person_list),
              len(company_telephone_list), len(company_address_list))
    count = 2
    for i in range(len(store_name_list)):
        key = 'A' + str(count)
        ali_excle[key] = store_name_list[i]
        key = 'B' + str(count)
        ali_excle[key] = english_name_list[i]
        key = 'C' + str(count)
        ali_excle[key] = store_url_list[i]
        key = 'D' + str(count)
        ali_excle[key] = main_product_list[i]
        key = 'E' + str(count)
        ali_excle[key] = year_list[i]
        key = 'F' + str(count)
        ali_excle[key] = main_country_list[i]
        key = 'G' + str(count)
        ali_excle[key] = money_list[i]
        key = 'H' + str(count)
        ali_excle[key] = manager_name_list[i]
        key = 'I' + str(count)
        ali_excle[key] = job_list[i]
        key = 'J' + str(count)
        ali_excle[key] = department_list[i]
        key = 'K' + str(count)
        ali_excle[key] = telephone_list[i]
        key = 'L' + str(count)
        ali_excle[key] = mobile_list[i]
        key = 'M' + str(count)
        ali_excle[key] = fax_list[i]
        count = count + 1
    save_name = 'ali_' + search_name + '.xls'
    wb2.save(save_name)
    driver.quit()

deep_crawl_main()