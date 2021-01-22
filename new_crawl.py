from selenium import webdriver
import time
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
from requests.exceptions import RequestException
import requests
from selenium.webdriver.support.ui import WebDriverWait


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


def crawl(url):
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
        print(url)
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


def get_name(driver):
    xpath = '//span[@title=\'Contacts\']'
    try:
        driver.find_element_by_xpath(xpath).click()
    except:
        time.sleep(3)
        try:
            driver.find_element_by_xpath(xpath).click()
        except:
            tup = ('None', 'None', 'None', 'None', 'None', 'None')
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
                print(e)
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
        name = 'None'
    department_div = soup.find('div', {'class': 'contact-department'})
    try:
        department = department_div.contents[0]
    except:
        department = 'None'
    job_div = soup.find('div', {'class': 'contact-job'})
    try:
        job = job_div.contents[0]
    except:
        job = 'None'
    table = str(soup.find('table', {'class': 'info-table'}))
    soup = BeautifulSoup(table, 'html.parser')
    trs = soup.find_all('tr')
    telephone = 'None'
    mobile = 'None'
    fax = 'None'
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


def log(driver):
    xpath = '//a[@title=\'Sign In\']'
    element = driver.find_element_by_xpath(xpath)
    driver.execute_script("arguments[0].click();", element)
    # ActionChains(driver).move_to_element(element).perform()
    # element.click()
    driver.implicitly_wait(20)
    xpath = '//input[@name=\'loginId\']'
    driver.find_element_by_xpath(xpath).send_keys('632080907@qq.com')
    xpath = '//input[@name=\'password\']'
    driver.find_element_by_xpath(xpath).send_keys('Ln632080907')
    driver.find_element_by_id('fm-login-submit').click()


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


if __name__ == '__main__':
    t1 = time.process_time()
    company_name = []
    url_set = []
    name_list = []
    department_list = []
    job_list = []
    telephone_list = []
    mobile_list = []
    fax_list = []
    wb = Workbook()
    excle = wb.active
    excle['A1'] = '名称'
    excle['B1'] = '连接'
    excle['C1'] = 'name'
    excle['D1'] = 'department'
    excle['E1'] = 'job'
    excle['F1'] = 'telephone'
    excle['G1'] = 'mobile'
    excle['H1'] = 'fax'
    # 把pen修改成你要搜索的商品名字
    search_name = 'casting'
    driver = webdriver.Chrome()
    driver.maximize_window()
    for num in range(1, 101):
        print(num)
        url = 'https://www.alibaba.com/products/' + search_name + '.html?IndexArea=product_en&page=' + str(num)
        driver.get(url)
        if num == 1:
            log(driver)
            pass
        js = "var q=document.documentElement.scrollTop=100000"
        driver.execute_script(js)
        time.sleep(3)
        html = driver.page_source
        try:
            soup = BeautifulSoup(html, 'html.parser')
        except:
            continue
        #                                                                        把list改成gallery
        a = soup.find_all('a', {'flasher-type': 'supplierName', 'class': 'organic-list-offer__seller-company'})
        for i in a:
            url = 'https:' + i.get('href')
            try:
                company = str(i.contents[0])
            except:
                continue
            if re.match(r'^(Shandong|Jinan|Qingdao|Zibo|Zaozhuang|Dongying|Yantai|Weifang|Jining|Taian'
                        r'|Weihai|Rizhao|Binzhou|Dezhou|Liaochen|Linxi|Heze)', company):
                print(url)
                tname = crawl(url)
                print(tname)
                if tname is None:
                    continue
                t_log(driver)
                print(url)
                driver.get(url)
                result = get_name(driver)
                print(result)
                name_list.append(result[0])
                department_list.append(result[1])
                job_list.append(result[2])
                telephone_list.append(result[3])
                mobile_list.append(result[4])
                fax_list.append(result[5])
                company_name.append(tname)
                url_set.append(url)

    lis = list(company_name)
    urls = list(url_set)
    count = 2
    stored_name = set()
    stored_url = set()
    for i in range(len(lis)):
        if lis[i] not in stored_name:
            # (name, department, job, telephone, mobile, fax)
            print(lis[i], urls[i])
            if '<td' in telephone_list[i]:
                driver.get(urls[i])
                result = get_name(driver)
                name_list[i] = result[0]
                department_list[i] = result[1]
                job_list[i] = result[2]
                telephone_list[i] = result[3]
                mobile_list[i] = result[4]
                fax_list[i] = result[5]
            stored_name.add(lis[i])
            stored_url.add(urls[i])
            key = 'A' + str(count)
            excle[key] = lis[i]
            key = 'B' + str(count)
            excle[key] = urls[i]
            key = 'C' + str(count)
            excle[key] = name_list[i]
            key = 'D' + str(count)
            excle[key] = department_list[i]
            key = 'E' + str(count)
            excle[key] = job_list[i]
            key = 'F' + str(count)
            excle[key] = telephone_list[i]
            key = 'G' + str(count)
            excle[key] = mobile_list[i]
            key = 'H' + str(count)
            excle[key] = fax_list[i]
            count = count + 1
        else:
            continue
    save_name = search_name + '.xls'
    wb.save(save_name)
    t2 = time.process_time()
    driver.quit()
    print(t2 - t1)