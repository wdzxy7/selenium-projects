from selenium import webdriver
import time
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook


def crawl(url):
    driver = webdriver.Chrome()
    driver.get(url)
    xpath = '//a[@data-spm=\'dassicon\']'
    try:
        driver.find_element_by_xpath(xpath).click()
    except:
        driver.refresh()
        time.sleep(4)
        driver.find_element_by_xpath(xpath).click()
    windows = driver.window_handles
    driver.switch_to.window(windows[-1])
    time.sleep(1)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find_all('div', {'class': 'ui2-balloon ui2-balloon-bl ui2-shadow-normal registration-name-tip'})
    div = str(div)
    name = re.findall(r'name "(.*?)" on', div)
    print(name[0])
    driver.quit()
    return name[0]


if __name__ == '__main__':
    company_name = set()
    wb = Workbook()
    excle = wb.active
    excle['A1'] = '名称'
    for num in range(1, 7):
        url = 'https://www.alibaba.com/products/casting.html?IndexArea=product_en&page=' + str(num)
        driver = webdriver.Chrome()
        driver.get(url)
        js = "var q=document.documentElement.scrollTop=100000"
        driver.execute_script(js)
        time.sleep(3)
        html = driver.page_source
        driver.quit()
        soup = BeautifulSoup(html, 'html.parser')
        a = soup.find_all('a', {'flasher-type': 'supplierName', 'class': 'organic-list-offer__seller-company'})
        for i in a:
            url = 'https:' + i.get('href')
            company = str(i.contents[0])
            if re.match(r'^(Shandong|Jinan|Qingdao|Zibo|Zaozhuang|Dongying|Yantai|Weifang|Jining|Taian'
                        r'|Weihai|Rizhao|Bingzhou|Dezhou|Liaochen|Linxi|Heze)', company):
                tname = crawl(url)
                company_name.add(tname)
                print(url)
    lis = list(company_name)
    count = 2
    for name in lis:
        key = 'A' + str(count)
        excle[key] = name
        count = count + 1
    wb.save("alibaba.xlsx")