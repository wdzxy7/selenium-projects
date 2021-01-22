import time
from selenium import webdriver
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook


wb = Workbook()
excle = wb.active
wb2 = Workbook()
excle2 = wb2.active
count = 2
count2 = 2


def creat_excle1():
    global excle
    excle['A1'] = '项目名称'
    excle['B1'] = '项目销售'
    excle['C1'] = '项目地块'
    excle['D1'] = '开盘时间'
    excle['E1'] = '入住时间'
    excle['F1'] = '平均价格'
    excle['G1'] = '城市'
    excle['H1'] = '总建筑面积'
    excle['I1'] = '区县/板块'
    excle['J1'] = '物业类型'
    excle['K1'] = '环线'
    excle['L1'] = '建筑类型'
    excle['M1'] = '项目特色'
    excle['N1'] = '装修情况'
    excle['O1'] = '开发商'
    excle['P1'] = '绿化率'
    excle['Q1'] = '容积率'
    excle['R1'] = '预售许可证'
    excle['S1'] = '项目地址'
    excle['T1'] = '分期所处地块'


def creat_excle2():
    global excle2
    excle2['A1'] = '城市'
    excle2['B1'] = '开发程度'
    excle2['C1'] = '总用地面积'
    excle2['D1'] = '建筑用地面积'
    excle2['E1'] = '征地面积'
    excle2['F1'] = '规划建筑面积'
    excle2['G1'] = '容积率'
    excle2['H1'] = '绿化率'
    excle2['I1'] = '商业比例'
    excle2['J1'] = '建筑密度'
    excle2['K1'] = '限制高度'
    excle2['L1'] = '配建保障房面积'
    excle2['M1'] = '出让方式'
    excle2['N1'] = '出让年限'
    excle2['O1'] = '配建保障房情况'
    excle2['P1'] = '地块所属分期'
    excle2['Q1'] = '公告时间'
    excle2['R1'] = '起始时间'
    excle2['S1'] = '截止时间'
    excle2['T1'] = '成交时间'
    excle2['U1'] = '公告编号'
    excle2['V1'] = '交易状况'
    excle2['W1'] = '起始价'
    excle2['X1'] = '成交价'
    excle2['Y1'] = '推出土地单价'
    excle2['Z1'] = '成交土地单价'
    excle2['AA1'] = '推出每亩价'
    excle2['AB1'] = '成交每亩价'
    excle2['AC1'] = '推出楼面价'
    excle2['AD1'] = '成交楼面价'
    excle2['AE1'] = '溢价率'
    excle2['AF1'] = '咨询电话'
    excle2['AG1'] = '保证金'
    excle2['AH1'] = '加价幅度'
    excle2['AI1'] = '出让单位'
    excle2['AJ1'] = '受让单位'
    excle2['AK1'] = '备注'
    excle2['AL1'] = '土地公告'


def get_stock(htmls):
    result = []
    for html in htmls:
        soup = BeautifulSoup(html, 'html.parser')
        a = soup.find_all('a')
        spans = soup.find_all('span')
        for name, number in zip(a, spans):
            tup = (name.contents[0], number.contents[0])
            result.append(tup)
    return result


def get_land_project(driver):
    xpath = '//span[text()=\'经营数据\']'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(0.5)
    xpath = '//span[text()=\'项目分布\']'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(0.5)
    xpath = '//a[text()=\'项目列表\']'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(3)
    company_number = 1
    project_number = 1
    while True:
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        tbody = soup.find('tbody')
        result = crawl_project_list(str(tbody))
        # 存放编码信息
        coding_list = []
        # 存放总信息
        store_list = []
        for lis in result:
            if len(lis) < 11:
                company_name = t_name
                project_name = lis[0]
                project_tup = (company_name, company_number, project_name, project_number)
                project_number = project_number + 1
                coding_list.append(project_tup)
                continue
            else:
                project_number = 1
                company_name = lis[1]
                project_name = lis[2]
                t_name = company_name
                project_tup = (company_name, company_number, project_name, project_number)
                company_number = company_number + 1
                coding_list.append(project_tup)
            try:
                driver.find_element_by_xpath('//a[text()=\'' + company_name + '\']').click()
            except:
                continue
            windows = driver.window_handles
            driver.switch_to.window(windows[-1])
            time.sleep(1)
            try:
                driver.find_element_by_xpath('//*[@id="aLeftMenuLink_deal"]/span[text()=\'项目销售\']').click()
            except:
                driver.close()
                driver.switch_to.window(windows[1])
                continue
            time.sleep(3)
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            tbody = soup.find('tbody')
            seal_message = crawl_project_list(str(tbody))
            driver.close()
            driver.switch_to.window(windows[1])
            t_list = []
            for seal in seal_message:
                if seal[0] == '合计':
                    continue
                seal.insert(0, project_name)
                seal.insert(0, company_name)
                store_list.append(tuple(seal))
        store1(store_list)
        store2(coding_list)
        try:
            time.sleep(1)
            driver.find_element_by_xpath('//ul[@id=\'listPager\']//font[text()=\'›\']').click()
            time.sleep(2)
        except:
            break


def crawl_project_list(html):
    soup = BeautifulSoup(html, 'html.parser')
    trs = soup.find_all('tr')
    return_list = []
    for tr in trs:
        soup = BeautifulSoup(str(tr), 'html.parser')
        tds = soup.find_all('td')
        t_list = []
        for td in tds:
            message = deal(str(td))
            t_list.append(message)
        return_list.append(t_list)
    return return_list


def deal(s):
    pattern = re.compile(r'<[^>]*>')
    remove = re.findall(pattern, s)
    for i in remove:
        s = s.replace(i, '').replace('\n', '').replace('\xa0', '')
    return s


def store1(result):
    global excle, count
    for lis in result:
        excle['A' + str(count)] = lis[0]
        excle['B' + str(count)] = lis[1]
        excle['C' + str(count)] = lis[2]
        excle['D' + str(count)] = lis[3]
        excle['E' + str(count)] = lis[4]
        excle['F' + str(count)] = lis[5]
        excle['G' + str(count)] = lis[6]
        excle['H' + str(count)] = lis[7]
        excle['I' + str(count)] = lis[8]
        excle['J' + str(count)] = lis[9]
        excle['K' + str(count)] = lis[10]
        excle['L' + str(count)] = lis[11]
        excle['M' + str(count)] = lis[12]
        excle['N' + str(count)] = lis[13]
        excle['O' + str(count)] = lis[14]
        excle['P' + str(count)] = lis[15]
        excle['Q' + str(count)] = lis[16]
        excle['R' + str(count)] = lis[17]
        excle['S' + str(count)] = lis[18]
        excle['T' + str(count)] = lis[19]
        count = count + 1


def store2(result):
    global excle2, count2
    for lis in result:
        excle2['A' + str(count2)] = lis[0]
        excle2['B' + str(count2)] = lis[1]
        excle2['C' + str(count2)] = lis[2]
        excle2['D' + str(count2)] = lis[3]
        count2 = count2 + 1


def get_fei(html):
    result = []
    soup = BeautifulSoup(html, 'html.parser')
    a = soup.find_all('a')
    for i in a:
        name = i.contents[0]
        num = '无'
        tup = (name, num)
        result.append(tup)
    return result


if __name__ == '__main__':
    creat_excle1()
    creat_excle2()
    url = 'https://creis.fang.com/'
    driver = webdriver.Chrome()
    driver.get(url)
    input()
    xpath = '/html/body/div[1]/div[1]/div[3]/a[2]'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(3)
    element = driver.find_element_by_id('tbodyHushi')
    hu_html = element.get_attribute('innerHTML')
    element = driver.find_element_by_id('tbodyShenshi')
    shen_html = element.get_attribute('innerHTML')
    element = driver.find_element_by_id('tbodyGangshi')
    gang_html = element.get_attribute('innerHTML')
    element = driver.find_element_by_id('tbodyFeishangshi')
    fei_html = element.get_attribute('innerHTML')
    htmls = []
    htmls.append(hu_html)
    htmls.append(shen_html)
    htmls.append(gang_html)
    # 获取所有股票信息
    stock_result = get_stock(htmls)
    place = '沪市'
    #fei = get_fei(fei_html)
    #stock_result = stock_result + fei
    temp = 1
    for i in stock_result:
        print(len(stock_result), temp, count)
        temp = temp + 1
        name = i[0]
        xpath = '//a[text()=\'' + name + '\']'
        try:
            driver.find_element_by_xpath(xpath).click()
        except:
            continue
        windows = driver.window_handles
        driver.switch_to.window(windows[-1])
        try:
            get_land_project(driver)
        except Exception as e:
            print(e)
            try:
                driver.close()
                driver.switch_to.window(windows[0])
            except:
                driver.switch_to.window(windows[0])
            continue
        driver.close()
        driver.switch_to.window(windows[0])
