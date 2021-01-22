import json
from selenium import webdriver
import time
from bs4 import BeautifulSoup
import xlrd
import xlutils.copy


class Qcc:
    def __init__(self, name):
        self.search_name = name
        # self.save_name = 'E://桌面文件//' + name + '.xls'
        self.save_name = name + '.xls'
        self.store_name_list = []
        self.company_telephone_list = []
        self.company_address_list = []
        self.company_legal_person_list = []
        self.driver = webdriver.Chrome()

    def star(self):
        path = 'ali_' + self.search_name + '.xls'
        data = xlrd.open_workbook(path)
        sheet = data.sheet_by_name("Sheet")
        # 复制表
        ws = xlutils.copy.copy(data)
        table = ws.get_sheet(0)
        # 获取搜索名称
        store_name_list = self.get_name(sheet)
        print(store_name_list)
        # 登录企查查
        self.driver.get('https://www.qcc.com/')
        for i in range(3):
            try:
                self.qcc_add_cookie()
                xpath = '//span[text()=\'登录 | 注册\']'
                try:
                    self.driver.find_element_by_xpath(xpath)
                    continue
                except:
                    break
            except:
                pass
        search = self.driver.find_element_by_id('searchkey')
        search.clear()
        search.send_keys('沂南县华闰天元机械有限公司')
        element = self.driver.find_element_by_class_name('index-searchbtn')
        self.driver.execute_script("arguments[0].click();", element)
        countnum = 0
        del store_name_list[0]
        # 开始搜索
        for name in store_name_list:
            print(name)
            if name is None:
                self.company_telephone_list.append('None')
                self.company_address_list.append('None')
                self.company_legal_person_list.append('None')
                continue
            search = self.driver.find_element_by_id('headerKey')
            search.clear()
            search.send_keys(name)
            time.sleep(1.5)
            self.driver.find_element_by_xpath('//button[text()=\'查一下\']').click()
            self.driver.find_element_by_class_name('ma_h1').click()
            windows = self.driver.window_handles
            self.driver.switch_to.window(windows[-1])
            html = self.driver.page_source
            time.sleep(1.8)
            self.driver.close()
            try:
                self.driver.switch_to.window(windows[0])
            except:
                self.open_page()
            result = self.get_telephon(html)
            self.company_telephone_list.append(result[0])
            self.company_address_list.append(result[1])
            self.company_legal_person_list.append(result[2])
            print(result)
            countnum += 1
            time.sleep(1.7)
            if countnum == 150:
                countnum = 0
                self.driver.quit()
                self.open_page()
                time.sleep(300)
        count = 1
        print(len(store_name_list),
              len(self.company_legal_person_list),
              len(self.company_telephone_list), len(self.company_address_list))
        # 保存
        for i in range(len(self.company_address_list)):
            legal_person = self.company_legal_person_list[i]
            telephone = self.company_telephone_list[i]
            address = self.company_address_list[i]
            table.write(count, 13, legal_person)
            table.write(count, 14, telephone)
            table.write(count, 15, address)
            count = count + 1

        ws.save(self.save_name)
        self.driver.quit()

    def qcc_add_cookie(self):
        try:
            xpath = '//*[@id="indexBonusModal"]/div/div/button'
            self.driver.find_element_by_xpath(xpath).click()
        except:
            pass
        with open('cookies(2).json', 'r', encoding='utf-8') as f:
            Cookies = json.loads(f.read())
        for cookie in Cookies:
            self.driver.add_cookie({
                "domain": cookie["domain"],
                "expirationDate": cookie['expirationDate'],
                "name": cookie['name'],
                "path": "/",
                "value": cookie['value'],
                "id": cookie['id']
            })
        # self.driver.get('https://www.qcc.com/')
        self.driver.refresh()

    def get_telephon(self, html):
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

    def get_name(self, sheet):
        names = []
        message_sum = sheet.nrows + 1
        for i in range(1, message_sum):
            name = sheet.cell_value(i - 1, 0)
            if name == '' or name == '名称':
                continue
            names.append(name)
        return names

    def open_page(self):
        self.driver = webdriver.Chrome()
        self.driver.get('https://www.qcc.com/')
        for i in range(3):
            try:
                self.qcc_add_cookie()
                xpath = '//span[text()=\'登录 | 注册\']'
                self.driver.find_element_by_xpath(xpath)
                break
            except:
                pass
        search = self.driver.find_element_by_id('searchkey')
        search.clear()
        search.send_keys('沂南县华闰天元机械有限公司')
        self.driver.find_element_by_class_name('index-searchbtn').click()

t_list = ['linyi sport', 'linyi grinding wheel']
for name in t_list:
    t = Qcc(name)
    t.star()