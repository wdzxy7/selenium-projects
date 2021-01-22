import time
from selenium import webdriver
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
import xlrd


wb = Workbook()
excle = wb.active
excle['A1'] = 'comment'
excle['B1'] = 'time'
excle['C1'] = 'country'
row = 2
country = ''


def log_in(driver):
    driver.implicitly_wait(10)
    xpath = '//input[@aria-label=\'手机号、帐号或邮箱\']'
    driver.find_element_by_xpath(xpath).send_keys('371269312@qq.com')
    xpath = '//input[@name=\'password\']'
    driver.find_element_by_xpath(xpath).send_keys('ahz2217995')
    xpath = '//div[@class=\'-MzZI\']/following-sibling::div[2]/button'
    driver.find_element_by_xpath(xpath).click()


def get_comment(driver):
    global excle, row, country
    time_list = []
    xpath = '//div[@class=\'zZYga\']//ul[contains(@class,"XQXOT")]'
    result = driver.find_element_by_xpath(xpath)
    times = driver.find_elements_by_tag_name('time')
    for t in times:
        time_list.append(t.get_attribute('title'))
    html = result.get_attribute('innerHTML')
    comment_list = crawl(html)
    for i in range(len(comment_list)):
        excle['A' + str(row)] = comment_list[i]
        excle['B' + str(row)] = time_list[i]
        excle['C' + str(row)] = country
        row = row + 1
        print(comment_list[i], time_list[i])
    xpath = '//div[@class=\'DdSX2\']//a[text()=\'下一页\']'
    try:
        driver.find_element_by_xpath(xpath).click()
        time.sleep(1)
        return 1
    except:
        return 0


def crawl(html):
    soup = BeautifulSoup(html, 'html.parser')
    return deep_crawl_button(str(soup))


def deep_crawl_button(html):
    pattern = re.compile(r'<span class=\"\">(.*?)</span>')
    result = re.findall(pattern, html)
    pattern = re.compile(r'<[^>]*>')
    for j in range(len(result)):
        remove = re.findall(pattern, result[j])
        for i in remove:
            result[j] = str(result[j]).replace(i, '')
    return result


def get_url(sheet):
    result = []
    country_list = []
    message_sum = sheet.nrows + 1
    for i in range(2, message_sum):
        name = sheet.cell_value(i-1, 1)
        country = sheet.cell_value(i-1, 0)
        country_list.append(country)
        result.append(name)
    return result, country_list


if __name__ == '__main__':
    path = 'country ins.xlsx'
    data = xlrd.open_workbook(path)
    sheet = data.sheet_by_name("Sheet1")
    cols = sheet.ncols + 1
    url_list, country_list = get_url(sheet)
    for i in range(len(url_list)):
        country = country_list[i]
        driver = webdriver.Chrome()
        try:
            driver.get(url_list[i])
        except:
            driver.refresh()
        log_in(driver)
        time.sleep(5)
        xpath = '//div[@class=\'eLAPa\']'
        driver.find_element_by_xpath(xpath).click()
        count = 0
        while count < 3:
            result = get_comment(driver)
            if result == 0:
                break
            print('**************************************************')
            count = count + 1
        '''
        while True:
            result = get_comment(driver)
            if result == 0:
                break
            print('**************************************************')
        '''
        driver.quit()
    wb.save("instagram.xlsx")