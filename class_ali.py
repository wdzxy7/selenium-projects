from selenium import webdriver
import time
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
from requests.exceptions import RequestException
import requests
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait


class Ali:
    def __init__(self, name):
        self.get_count = 0
        self.search_name = name
        self.save_name = 'ali_' + name + '.xls'
        self.driver = webdriver.Chrome()
        self.main_country_list = []
        self.store_url_list = []
        self.store_name_list = []
        self.year_list = []
        self.english_name_list = []
        self.chinese_name_list = []
        self.main_product_list = []
        self.sold_list = []
        self.money_list = []
        self.url_list = []
        self.manager_name_list = []
        self.department_list = []
        self.job_list = []
        self.telephone_list = []
        self.mobile_list = []
        self.fax_list = []
        self.stored_url = []
        self.compare_list = ['Shandong', 'Jinan', 'Qingdao', 'Zibo', 'Zaozhuang', 'Dongying', 'Yantai', 'Weifang',
                             'Jining', 'Taian', 'Weihai', 'Rizhao', 'Binzhou', 'Dezhou', 'Liaocheng', 'Linyi', 'Heze']

    class Warn(Exception):
        pass

    def deep_crawl_main(self):
        wb2 = Workbook()
        ali_excle = wb2.active
        self.creat_ali_wb(ali_excle)
        self.driver.maximize_window()
        url = 'https://www.alibaba.com/trade/search?fsb=y&IndexArea=company_en&CatId=&SearchText=' + self.search_name
        self.driver.get(url)
        self.log()
        while True:
            js = "var q=document.documentElement.scrollTop=100000"
            self.driver.execute_script(js)
            xpath = '//div[@class=\'grid grid-c3-s5e6\']'
            res = self.driver.find_element_by_xpath(xpath)
            html = res.get_attribute('innerHTML')
            if html is None:
                continue
            self.crawl(html)
            xpath = '//a[@class=\'next\']'
            try:
                next_button = self.driver.find_element_by_xpath(xpath)
            except:
                break
            next_button.click()
            time.sleep(1)
        self.driver.quit()
        print(len(self.store_name_list), len(self.store_url_list), len(self.chinese_name_list), len(self.main_product_list), len(self.year_list),
              len(self.main_country_list),
              len(self.money_list), len(self.manager_name_list)
              , len(self.job_list), len(self.department_list), len(self.telephone_list), len(self.mobile_list), len(self.fax_list),)
        self.driver = webdriver.Chrome()
        self.driver.get(url)
        self.log()
        time.sleep(4)
        stored_name = set()
        search_sum = 0
        length = len(self.chinese_name_list)
        del_count = 0
        for i in range(length):
            if self.chinese_name_list[i] not in stored_name:
                company_name = self.chinese_name_list[i]
                print(self.chinese_name_list[i], self.url_list[i])
                stored_name.add(self.chinese_name_list[i])
                url = self.url_list[i]
                try:
                    self.driver.get(url)
                except:
                    result = ['', '', '', '', '', '']
                    self.store_name_list.append(company_name)
                    self.store_url_list.append(url)
                    self.manager_name_list.append(result[0])
                    self.department_list.append(result[1])
                    self.job_list.append(result[2])
                    self.telephone_list.append(result[3])
                    self.mobile_list.append(result[4])
                    self.fax_list.append(result[5])
                    continue
                result = self.get_name()
                print(result)
                self.store_name_list.append(company_name)
                self.store_url_list.append(url)
                self.manager_name_list.append(result[0])
                self.department_list.append(result[1])
                self.job_list.append(result[2])
                self.telephone_list.append(result[3])
                self.mobile_list.append(result[4])
                self.fax_list.append(result[5])
                time.sleep(0.5)
                # 查询多少次检测一下是否还处于登录状态
                if search_sum == 10:
                    self.driver.quit()
                    self.driver = webdriver.Chrome()
                    search_url = 'https://www.alibaba.com/trade/search?fsb=y&IndexArea=company_en&CatId=&SearchText=' + self.search_name
                    self.driver.get(search_url)
                    try:
                        self.driver.implicitly_wait(10)
                        self.log()
                        time.sleep(3)
                    except:
                        pass
                    search_sum = 0
                search_sum = search_sum + 1
                self.get_count = self.get_count + 1
                # 爬取多少份暂停5分钟
                if self.get_count > 100:
                    self.driver.quit()
                    self.driver = webdriver.Chrome()
                    search_url = 'https://www.alibaba.com/trade/search?fsb=y&IndexArea=company_en&CatId=&SearchText=' + self.search_name
                    self.driver.get(search_url)
                    try:
                        self.driver.implicitly_wait(10)
                        self.log()
                        time.sleep(3)
                    except:
                        pass
                    self.get_count = 0
                    # 爬取100份后暂停时间
                    time.sleep(300)
            else:
                del self.english_name_list[i - del_count]
                del self.main_country_list[i - del_count]
                del self.main_product_list[i - del_count]
                del self.year_list[i - del_count]
                del self.money_list[i - del_count]
                del_count = del_count + 1
        count = 2
        for i in range(len(self.store_name_list)):
            key = 'A' + str(count)
            ali_excle[key] = self.store_name_list[i]
            key = 'B' + str(count)
            ali_excle[key] = self.english_name_list[i]
            key = 'C' + str(count)
            ali_excle[key] = self.store_url_list[i]
            key = 'D' + str(count)
            ali_excle[key] = self.main_product_list[i]
            key = 'E' + str(count)
            ali_excle[key] = self.year_list[i]
            key = 'F' + str(count)
            ali_excle[key] = self.main_country_list[i]
            key = 'G' + str(count)
            ali_excle[key] = self.money_list[i]
            key = 'H' + str(count)
            ali_excle[key] = self.manager_name_list[i]
            key = 'I' + str(count)
            ali_excle[key] = self.job_list[i]
            key = 'J' + str(count)
            ali_excle[key] = self.department_list[i]
            key = 'K' + str(count)
            ali_excle[key] = self.telephone_list[i]
            key = 'L' + str(count)
            ali_excle[key] = self.mobile_list[i]
            key = 'M' + str(count)
            ali_excle[key] = self.fax_list[i]
            count = count + 1

        wb2.save(self.save_name)
        self.driver.quit()

    def get_html(self, url):
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

    def log(self):
        xpath = '//a[@data-val=\'ma_signin\']'
        element = self.driver.find_element_by_xpath(xpath)
        self.driver.execute_script("arguments[0].click();", element)
        # ActionChains(driver).move_to_element(element).perform()
        # element.click()
        self.driver.implicitly_wait(20)
        xpath = '//input[@name=\'loginId\']'
        self.driver.find_element_by_xpath(xpath).send_keys('632080907@qq.com')
        xpath = '//input[@name=\'password\']'
        self.driver.find_element_by_xpath(xpath).send_keys('Ln632080907')
        xpath = '//span[@id=\'nc_1_n1z\']'
        try:
            span = self.driver.find_element_by_xpath(xpath)
            action = ActionChains(self.driver)
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
        self.driver.find_element_by_id('fm-login-submit').click()

    def crawl(self, html):
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
            english_name = self.deal(str(h2)).replace('\n', '')
            try:
                div = soup.find('div', {'class': 'value ellipsis ph'})
                main_product = self.deal(str(div))
            except:
                main_product = 'None'
            try:
                div = soup.find('div', {'class': 'lab'})
                sold = self.deal(str(div)).replace(' ', '').replace('\n', '')
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
            temp_name = english_name.split(' ', 1)
            #if re.match(r'(Shandong|Jinan|Qingdao|Zibo|Zaozhuang|Dongying|Yantai|Weifang|Jining|Taian'
                        #r'|Weihai|Rizhao|Binzhou|Dezhou|Liaocheng|Linyi|Heze)', temp_name[0]):
            if temp_name[0] in self.compare_list:
                chinese_name = self.crawl_name(url)
                self.english_name_list.append(english_name)
                self.chinese_name_list.append(chinese_name)
                self.year_list.append(year)
                self.main_product_list.append(main_product)
                self.sold_list.append(sold)
                self.money_list.append(money)
                self.url_list.append(url)
                self.main_country_list.append(str(country))
                # print(url)
                print(year, end='\n****************\n')
                print(english_name, end='\n***************\n')
                # print(chinese_name)
                # print(country, end='\n*****************\n')
                # print(main_product, end='\n*****************\n')
                print(sold, end='\n***************************\n')
                print(money, end='\n*************************\n')
                print('-----------------------------------------------------------------------------------')

    def deal(self, s):
        pattern = re.compile(r'<[^>]*>')
        remove = re.findall(pattern, s)
        for i in remove:
            s = s.replace(i, '')
        return s

    def crawl_name(self, url):
        try:
            html = self.get_html(url)
            soup = BeautifulSoup(html, 'html.parser')
        except:
            return None
        try:
            a = soup.find_all('a', {'data-spm': 'dassicon'})
            if len(a) == 0:
                raise self.Warn()
        except self.Warn:
            try:
                a = soup.find_all('a', {'data-spm': 'davicon'})
                if len(a) == 0:
                    raise self.Warn()
            except self.Warn:
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
        html = self.get_html(url)
        soup = BeautifulSoup(html, 'html.parser')
        div = soup.find_all('div', {'class': 'ui2-balloon ui2-balloon-bl ui2-shadow-normal registration-name-tip'})
        div = str(div)
        name = re.findall(r'name "(.*?)" on', div)
        if len(name) == 0:
            return None
        return name[0]

    def get_name(self):
        xpath = '//span[@title=\'Contacts\']'
        try:
            self.driver.find_element_by_xpath(xpath).click()
        except:
            time.sleep(3)
            try:
                self.driver.find_element_by_xpath(xpath).click()
            except:
                tup = ('', '', '', '', '', '')
                return tup

        xpath = '//a[text()=\'View details\']'
        wait = WebDriverWait(self.driver, 4, 0.2)
        try:
            wait.until(lambda dribver: self.driver.find_element_by_xpath('//a[text()=\'View details\']'))
            for i in range(40):
                try:
                    element = self.driver.find_element_by_xpath(xpath)
                    self.driver.execute_script("arguments[0].click();", element)
                    break
                except Exception as e:
                    pass
        except:
            pass
        time.sleep(1.5)
        html = self.driver.page_source
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

    def creat_ali_wb(self, excle):
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

