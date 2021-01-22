# coding=utf-8
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup
import io
import xlwt
from selenium.webdriver.common.keys import Keys
# from fake_useragent import UserAgent
import sys
import time
import xlrd
import datetime

chrome_options = Options()
prefs = {'profile.managed_default_content_settings.images': 2}
url = 'http://sia.ipsos.cn/huawei/#/login'
driver = webdriver.Chrome(options=chrome_options)
driver.maximize_window()
driver.set_page_load_timeout(20)
try:
    driver.get(url)
except:
    driver.execute_script('window.stop()')


def get_time():
    year = datetime.date.today()
    result = str(year).split('-')
    result1 =str(year).split('-')
    big = {1, 3, 5, 7, 8, 10, 12}
    mini = {4, 6, 9, 11}
    if int(result[2]) == 2:
        month = int(result[1]) - 1
        if month == 2:
            if (int(result[0]) % 4 == 0 and int(result[0]) % 100 != 0) or (int(result[0]) % 400 == 0):
                if month < 10:
                    result[1] = '0' + str(month)
                else:
                    result[1] = str(month)
                result[2] = '29'
            else:
                if month < 10:
                    result[1] = '0' + str(month)
                else:
                    result[1] = str(month)
                result[2] = '28'
        elif month in big:
            if month < 10:
                result[1] = '0' + str(month)
            else:
                result[1] = str(month)
            result[2] = '31'
        elif month in mini:
            if month < 10:
                result[1] = '0' + str(month)
            else:
                result[1] = str(month)
            result[2] = '30'
        else:
            result[0] = str(int(result[0]) - 1)
            result[1] = '12'
            result[2] = '31'
    elif int(result[2]) == 1:
        month = int(result[1]) - 1
        if month == 2:
            if (int(result[0]) % 4 == 0 and int(result[0]) % 100 != 0) or (int(result[0]) % 400 == 0):
                if month < 10:
                    result[1] = '0' + str(month)
                else:
                    result[1] = str(month)
                result[2] = '28'
            else:
                result[2] = '27'
                if month < 10:
                    result[1] = '0' + str(month)
                else:
                    result[1] = str(month)
        elif month in big:
            result[2] = '30'
            if month < 10:
                result[1] = '0' + str(month)
            else:
                result[1] = str(month)
        elif month in mini:
            result[1] = str(month)
            result[2] = '29'
        else:
            result[0] = str(int(result[0]) - 1)
            result[1] = '12'
            result[2] = '30'
    else:
        result[2] = str(int(result[2]) - 2)
    times = result[0] + '-' + result[1] + '-' + result[2] + ' - ' + result1[0] + '-' + result1[1] + '-' + result1[2]
    return times

def find_the_bottom(a):
    global action
    i=0
    while i<30:
        i+=1
        try:
            driver.find_element_by_xpath(a).click()
            print(i)
            i=100
        except:
            time.sleep(0.5)


def click_locxy(dr, x, y, left_click=True):
    if left_click:
        ActionChains(dr).move_by_offset(x, y).click().perform()
    else:
        ActionChains(dr).move_by_offset(x, y).context_click().perform()
    ActionChains(dr).move_by_offset(-x, -y).perform()


def new_find_the_bottom(a):
    global action
    i=0
    while i<30:
        i+=1
        try:
            # action.move_to_element(a).click()
            driver.find_element_by_xpath(a).click()
            print(i)
            i=100
        except:
            time.sleep(3)
# def find_the_bottom(element_value,element_type,retry_time):
#     i=0
#     while i<retry_time:
#         i+=1
#         try:
#             if element_type == 'id':
#                 driver.find_element_by_id(element_value).click()
# 			if element_type == 'class':
# 				driver.find_element_by_id(element_value).click()
# 			if element_type == 'xPath':
#                 driver.find_element_by_id(element_value).click()
#             i=100000
#         except:
#             time.sleep(2)

find_the_bottom('//*[@id="app"]/div/div/div[4]/div/div')
#登录
driver.find_element_by_xpath('//input[@placeholder=\"Please enter your email\"]').send_keys('hihonordb@huawei.com')
driver.find_element_by_xpath('//input[@placeholder=\"Please enter the password\"]').send_keys('123456')
find_the_bottom('//div[@class=\"loginIn-subn-inputWitch\"]')
find_the_bottom('//div[@class=\"logins-content-path\"]')
print("登录")
#选择Brand
find_the_bottom('//input[@class=\"ivu-select-input\"]')
find_the_bottom('//*[@class=\'ivu-select-dropdown-list\'][contains(text(),"Honor Brand")]')
find_the_bottom('//*[@id="app"]/div/div/div[2]/div/div/div[1]/div[1]/div[1]/div[1]/div/div[2]/ul[2]/li[13]')
time.sleep(3)
print('brand')


#选择Topic
action=ActionChains(driver)
a=driver.find_element_by_xpath('//*[@name="tree1"]/div[1]/div[1]/div')
action.move_to_element(a).click().perform()
find_the_bottom('//*[@name="tree1"]/div[1]/div[1]/div')#点击topic
find_the_bottom('//*[@x-placement="bottom-start"][@class="ivu-select-dropdown linkagepage"]/div[5]/div/div[2][contains(text(),"Honor Brand")]')#点击Honor Brand
find_the_bottom('//*[@class="input-span"][contains(text(),"Topic")]/../div/div[2]/div[2]/div[3]')#点击Confirm
time.sleep(30)
print('topic')
# #选择More
d=driver.find_element_by_xpath('//*[@class="TomHeader-top-cust"]/div[2]')
# action.move_to_element(d).click()
new_find_the_bottom('//*[@class="TomHeader-top-cust"]/div[2]')
print('more')

#选择时间
b=driver.find_element_by_xpath('//*[@name="datepicker4"]/div/div/div/input[1]')
action.move_to_element(b).click()
find_the_bottom('//*[@name="datepicker4"]/div/div/div/input[1]')
find_the_bottom('//*[@x-placement="bottom-end"]/div/div/div/div')
find_the_bottom('//*[@name="datepicker4"]/div/div[2]/div/div/div[2]/div[4]/button[2]')

#下载
find_the_bottom('//img[@class=\'download\']')

time.sleep(5)
xpath = '//input[@placeholder=\'Enter the starting location to download\']'
result = driver.find_elements_by_xpath(xpath)
result[0].send_keys('1')
result[1].send_keys('50000')

xpath = '//div[@class=\'download_frame\']/footer/button/following-sibling::button[1]'
driver.find_element_by_xpath(xpath).click()