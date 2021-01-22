from selenium import webdriver
import time
from bs4 import BeautifulSoup
from openpyxl import Workbook

wb = Workbook()
excle = wb.active
excle['A1'] = '风险等级'
excle['B1'] = '性别'
excle['C1'] = '年龄'
excle['D1'] = '婚姻'
excle['E1'] = '收入'
excle['F1'] = '公司行业'
excle['G1'] = '其他负债'
excle['H1'] = '申请借款'
excle['I1'] = '逾期次数'
count = 2


def login(driver):
    xpath = '//a[@href=\"/login\"]'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(2)
    xpath = '//span[@class=\"tab-password\"]'
    driver.find_element_by_xpath(xpath).click()
    xpath = '//input[@placeholder=\"手机号/邮箱\"]'
    username = driver.find_element_by_xpath(xpath)
    username.send_keys('15391616023')
    xpath = '//input[@placeholder=\"密码\"]'
    password = driver.find_element_by_xpath(xpath)
    password.send_keys('nfl123456')
    xpath = '//span[@id=\"rememberme-login\"]'
    driver.find_element_by_xpath(xpath).click()
    xpath = '//button[@class=\"is-allow\"]'
    driver.find_element_by_xpath(xpath).click()


def crawl(html):
    soup = BeautifulSoup(html, 'html.parser')
    span = soup.find_all('span', {'data-reactid': '.0.1.0.2.$3.1'})
    for i in span:
        for string in i.stripped_strings:
            level = string
    em = soup.find_all('em', {'data-reactid': '.2.1.0.0.1.0.4.1'})
    for i in em:
        for string in i.stripped_strings:
            sex = string
    em = soup.find_all('em', {'data-reactid': '.2.1.0.0.1.0.6.1'})
    for i in em:
        for string in i.stripped_strings:
            age = string
    em = soup.find_all('em', {'data-reactid': '.2.1.0.0.1.0.8.1'})
    for i in em:
        for string in i.stripped_strings:
            marriage = string
    em = soup.find_all('em', {'data-reactid': '.2.1.0.0.1.0.9.1'})
    for i in em:
        for string in i.stripped_strings:
            income = string
    em = soup.find_all('em', {'data-reactid': '.2.1.0.0.1.0.e.1'})
    for i in em:
        for string in i.stripped_strings:
            industry = string
    em = soup.find_all('em', {'data-reactid': '.2.1.0.0.1.0.j.1'})
    for i in em:
        for string in i.stripped_strings:
            debt = string
    span = soup.find_all('span', {'data-reactid': '.2.1.0.0.1.2.0.1.0'})
    for i in span:
        for string in i.stripped_strings:
            times = string
    span = soup.find_all('span', {'data-reactid': '.2.1.0.0.1.2.5.1.0'})
    for i in span:

        for string in i.stripped_strings:
            time_out = string
    print(level, sex, age, marriage, income, industry, debt, times, time_out)
    store(level, sex, age, marriage, income, industry, debt, times, time_out)


def store(level, sex, age, marriage, income, industry, debt, times, time_out):
    global count, excle
    excle['A' + str(count)] = level
    excle['B' + str(count)] = sex
    excle['C' + str(count)] = age
    excle['D' + str(count)] = marriage
    excle['E' + str(count)] = income
    excle['F' + str(count)] = industry
    excle['G' + str(count)] = debt
    excle['H' + str(count)] = times
    excle['I' + str(count)] = time_out
    count = count + 1


if __name__ == '__main__':
    url = 'https://www.renrendai.com/loan.html'
    driver = webdriver.Chrome()
    driver.get(url)
    login(driver)
    driver.refresh()
    for i in range(0, 10):
        data = '.0.0.1:$' + str(i) + '.0.1.0.2'
        xpath = '//a[@data-reactid=\"' + data + '\"]'
        time.sleep(2)
        try:
            driver.find_element_by_xpath(xpath).click()
            windows = driver.window_handles
            driver.switch_to.window(windows[-1])
            time.sleep(1)
            html = driver.page_source
            crawl(html)
            driver.switch_to.window(windows[0])
        except :
            break
    driver.quit()
wb.save("loan.xlsx")