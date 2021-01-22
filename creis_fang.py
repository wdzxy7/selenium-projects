import time
from selenium import webdriver
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook


wb = Workbook()
excle = wb.active
count = 2
wb2 = Workbook()
excle2 = wb2.active
count2 = 2
company_number = 23008


def creat_excle1():
    global excle
    excle['A1'] = '城市'
    excle['B1'] = '总期编码'
    excle['C1'] = '总期名称'
    excle['D1'] = '分期编码'
    excle['E1'] = '分期名称'
    excle['F1'] = '物业类型'
    excle['G1'] = '销售状态'
    excle['H1'] = '环线'
    excle['I1'] = '区县/板块'
    excle['J1'] = '项目地址'
    excle['K1'] = '权益'
    excle['L1'] = '规划建筑面积（万m2）'
    excle['M1'] = '装修情况'
    excle['N1'] = '容积率'
    excle['O1'] = '开盘时间'
    excle['P1'] = '户数'
    excle['Q1'] = '报价'
    excle['R1'] = '开发商'
    excle['S1'] = '地块名称'
    excle['T1'] = '楼面价'
    excle['U1'] = '拿地时间'
    excle['V1'] = '项目销售'
    excle['W1'] = '项目地块'
    excle['X1'] = '开盘时间'
    excle['Y1'] = '入住时间'
    excle['Z1'] = '平均价格'
    excle['AA1'] = '城市'
    excle['AB1'] = '总建筑面积'
    excle['AC1'] = '区县/板块'
    excle['AD1'] = '物业类型'
    excle['AE1'] = '环线'
    excle['AF1'] = '建筑类型'
    excle['AG1'] = '项目特色'
    excle['AH1'] = '装修情况'
    excle['AI1'] = '开发商'
    excle['AJ1'] = '绿化率'
    excle['AK1'] = '容积率'
    excle['AL1'] = '预售许可证'
    excle['AM1'] = '项目地址'
    excle['AN1'] = '分期所处地块'
    excle['AO1'] = '城市'
    excle['AP1'] = '开发程度'
    excle['AQ1'] = '总用地面积'
    excle['AR1'] = '建筑用地面积'
    excle['AS1'] = '征地面积'
    excle['AT1'] = '规划建筑面积'
    excle['AU1'] = '容积率'
    excle['AV1'] = '绿化率'
    excle['AW1'] = '商业比例'
    excle['AX1'] = '建筑密度'
    excle['AY1'] = '限制高度'
    excle['AZ1'] = '配建保障房面积'
    excle['BA1'] = '出让方式'
    excle['BB1'] = '出让年限'
    excle['BC1'] = '配建保障房情况'
    excle['BD1'] = '地块所属分期'
    excle['BE1'] = '公告时间'
    excle['BF1'] = '起始时间'
    excle['BG1'] = '截止时间'
    excle['BH1'] = '成交时间'
    excle['BI1'] = '公告编号'
    excle['BJ1'] = '交易状况'
    excle['BK1'] = '起始价'
    excle['BL1'] = '成交价'
    excle['BM1'] = '推出土地单价'
    excle['BN1'] = '成交土地单价'
    excle['BO1'] = '推出每亩价'
    excle['BP1'] = '成交每亩价'
    excle['BQ1'] = '推出楼面价'
    excle['BR1'] = '成交楼面价'
    excle['BS1'] = '溢价率'
    excle['BT1'] = '咨询电话'
    excle['BU1'] = '保证金'
    excle['BV1'] = '加价幅度'
    excle['BW1'] = '出让单位'
    excle['BX1'] = '受让单位'
    excle['BY1'] = '备注'
    excle['BZ1'] = '土地公告'


def creat_excle2():
    global excle2
    excle2['A1'] = '项目总期'
    excle2['B1'] = '总期编号'
    excle2['C1'] = '时间'
    excle2['D1'] = '成交套数'
    excle2['E1'] = '成交面积(㎡)'
    excle2['F1'] = '成交单价(元/㎡)'
    excle2['G1'] = '成交金额(万元)'


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
    global company_number
    xpath = '//span[text()=\'经营数据\']'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(0.5)
    xpath = '//span[text()=\'项目分布\']'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(0.5)
    xpath = '//a[text()=\'项目列表\']'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(1)
    xpath = '//span[text()=\'指标选择\']'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(1)
    for i in range(3):
        driver.find_element_by_xpath('//*[@id="indexSetContent"]//a[text()=\'全选\']').click()
    time.sleep(0.5)
    for i in range(30):
        try:
            driver.find_element_by_xpath('//a[@id=\'but_confirm\']').click()
        except:
            break
    time.sleep(2)
    project_number = 1
    while True:
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        tbody = soup.find('tbody')
        result = crawl_project_list(str(tbody))
        # 存放项目分期信息
        project_staging = []
        city = ''
        company = ''
        t_company = ''
        store_seal_list = []
        for lis in result:
            if len(lis) < 25:
                project_number = project_number + 1
                company_name = lis[0]
                t_lis = lis
                t_lis.insert(0, company)
                t_lis.insert(0, city)
                t_lis.insert(1, company_number)
                t_lis.insert(3, project_number)
            else:
                project_number = 1
                city = lis[0]
                company = lis[1]
                company_name = lis[2]
                if t_company != company:
                    t_company = company
                    company_number = company_number + 1
                t_lis = lis
                t_lis.insert(1, company_number)
                t_lis.insert(3, project_number)
            try:
                driver.find_element_by_xpath('//a[text()=\'' + company_name + '\']').click()
            except:
                continue
            windows = driver.window_handles
            driver.switch_to.window(windows[-1])
            time.sleep(3.5)
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            tbody = soup.find('tbody')
            land_message = crawl_project_list(str(tbody))
            try:
                driver.find_element_by_xpath('//*[@id="aLeftMenuLink_deal"]/span[text()=\'项目销售\']').click()
                time.sleep(4)
                html = driver.page_source
                soup = BeautifulSoup(html, 'html.parser')
                tbody = soup.find('tbody')
                seal_message = crawl_project_list(str(tbody))
                for seal in seal_message:
                    if seal[0] == '合计':
                        continue
                    seal.insert(0, company)
                    seal.insert(1, company_number)
                    store_seal_list.append(tuple(seal))
                pdd = 1
            except:
                pdd = 0
            try:
                driver.find_element_by_xpath('//*[@id="aLeftMenuLink_land"]/span[text()=\'项目地块\']').click()
                time.sleep(2)
                html = driver.page_source
                soup = BeautifulSoup(html, 'html.parser')
                tbody = soup.find_all('tbody')
                land1 = crawl_project_list(str(tbody[0]))
                land2 = crawl_project_list(str(tbody[1]))
                lands = land1 + land2
                second_list = []
                for land in lands:
                    for number in range(1, len(land), 2):
                        second_list.append(land[number])
                pdd2 = 1
            except:
                pdd2 = 0
                second_list = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
            tup = (pdd, pdd2)
            tup = list(tup)
            driver.close()
            driver.switch_to.window(windows[1])
            third_list = []
            t_count = 0
            for land in land_message:
                for number in range(1, len(land), 2):
                    if t_count == 17:
                        break
                    third_list.append(land[number])
                    t_count = t_count + 1
            third_list = tup + third_list
            final_list = t_lis + third_list + second_list
            project_staging.append(tuple(final_list))
        store1(project_staging)
        store2(store_seal_list)
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
        print(len(lis), lis)
        if len(lis) == 84:
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
            excle['R' + str(count)] = lis[23]
            excle['S' + str(count)] = lis[24]
            excle['T' + str(count)] = lis[25]
            excle['U' + str(count)] = lis[26]
            excle['V' + str(count)] = lis[27]
            excle['W' + str(count)] = lis[28]
            excle['X' + str(count)] = lis[29]
            excle['Y' + str(count)] = lis[30]
            excle['Z' + str(count)] = lis[31]
            excle['AA' + str(count)] = lis[32]
            excle['AB' + str(count)] = lis[33]
            excle['AC' + str(count)] = lis[34]
            excle['AD' + str(count)] = lis[35]
            excle['AE' + str(count)] = lis[36]
            excle['AF' + str(count)] = lis[37]
            excle['AG' + str(count)] = lis[38]
            excle['AH' + str(count)] = lis[39]
            excle['AR' + str(count)] = lis[40]
            excle['AJ' + str(count)] = lis[41]
            excle['AK' + str(count)] = lis[42]
            excle['AL' + str(count)] = lis[43]
            excle['AM' + str(count)] = lis[44]
            excle['AN' + str(count)] = lis[45]
            excle['AO' + str(count)] = lis[46]
            excle['AP' + str(count)] = lis[47]
            excle['AQ' + str(count)] = lis[48]
            excle['AR' + str(count)] = lis[49]
            excle['AS' + str(count)] = lis[50]
            excle['AT' + str(count)] = lis[51]
            excle['AU' + str(count)] = lis[52]
            excle['AV' + str(count)] = lis[53]
            excle['AW' + str(count)] = lis[54]
            excle['AX' + str(count)] = lis[55]
            excle['AY' + str(count)] = lis[56]
            excle['AZ' + str(count)] = lis[57]
            excle['BA1' + str(count)] = lis[58]
            excle['BB1' + str(count)] = lis[59]
            excle['BC1' + str(count)] = lis[60]
            excle['BD1' + str(count)] = lis[61]
            excle['BE1' + str(count)] = lis[62]
            excle['BF1' + str(count)] = lis[63]
            excle['BG1' + str(count)] = lis[64]
            excle['BH1' + str(count)] = lis[65]
            excle['BI1' + str(count)] = lis[66]
            excle['BJ1' + str(count)] = lis[67]
            excle['BK1' + str(count)] = lis[68]
            excle['BL1' + str(count)] = lis[69]
            excle['BM1' + str(count)] = lis[70]
            excle['BN1' + str(count)] = lis[71]
            excle['BO1' + str(count)] = lis[72]
            excle['BP1' + str(count)] = lis[73]
            excle['BQ1' + str(count)] = lis[74]
            excle['BR1' + str(count)] = lis[75]
            excle['BS1' + str(count)] = lis[76]
            excle['BT1' + str(count)] = lis[77]
            excle['BU1' + str(count)] = lis[78]
            excle['BV1' + str(count)] = lis[79]
            excle['BW1' + str(count)] = lis[80]
            excle['BX1' + str(count)] = lis[81]
            excle['BY1' + str(count)] = lis[82]
            excle['BZ1' + str(count)] = lis[83]
            count = count + 1
        elif len(lis) == 63:
            excle['A' + str(count)] = lis[0]
            excle['B' + str(count)] = lis[1]
            excle['C' + str(count)] = lis[2]
            excle['D' + str(count)] = lis[3]
            excle['E' + str(count)] = lis[4]
            excle['F' + str(count)] = ''
            excle['G' + str(count)] = ''
            excle['H' + str(count)] = ''
            excle['I' + str(count)] = ''
            excle['J' + str(count)] = ''
            excle['K' + str(count)] = ''
            excle['L' + str(count)] = ''
            excle['M' + str(count)] = ''
            excle['N' + str(count)] = ''
            excle['O' + str(count)] = ''
            excle['P' + str(count)] = ''
            excle['Q' + str(count)] = ''
            excle['R' + str(count)] = ''
            excle['S' + str(count)] = ''
            excle['T' + str(count)] = ''
            excle['U' + str(count)] = lis[5]
            excle['V' + str(count)] = lis[6]
            excle['W' + str(count)] = lis[7]
            excle['X' + str(count)] = lis[8]
            excle['Y' + str(count)] = lis[9]
            excle['Z' + str(count)] = lis[10]
            excle['AA' + str(count)] = lis[11]
            excle['AB' + str(count)] = lis[12]
            excle['AC' + str(count)] = lis[13]
            excle['AD' + str(count)] = lis[14]
            excle['AE' + str(count)] = lis[15]
            excle['AF' + str(count)] = lis[16]
            excle['AG' + str(count)] = lis[17]
            excle['AH' + str(count)] = lis[18]
            excle['AR' + str(count)] = lis[19]
            excle['AJ' + str(count)] = lis[20]
            excle['AK' + str(count)] = lis[21]
            excle['AL' + str(count)] = lis[22]
            excle['AM' + str(count)] = lis[23]
            excle['AN' + str(count)] = lis[24]
            excle['AO' + str(count)] = lis[25]
            excle['AP' + str(count)] = lis[26]
            excle['AQ' + str(count)] = lis[27]
            excle['AR' + str(count)] = lis[28]
            excle['AS' + str(count)] = lis[29]
            excle['AT' + str(count)] = lis[30]
            excle['AU' + str(count)] = lis[31]
            excle['AV' + str(count)] = lis[32]
            excle['AW' + str(count)] = lis[33]
            excle['AX' + str(count)] = lis[34]
            excle['AY' + str(count)] = lis[35]
            excle['AZ' + str(count)] = lis[36]
            excle['BA1' + str(count)] = lis[37]
            excle['BB1' + str(count)] = lis[38]
            excle['BC1' + str(count)] = lis[39]
            excle['BD1' + str(count)] = lis[40]
            excle['BE1' + str(count)] = lis[41]
            excle['BF1' + str(count)] = lis[42]
            excle['BG1' + str(count)] = lis[43]
            excle['BH1' + str(count)] = lis[44]
            excle['BI1' + str(count)] = lis[45]
            excle['BJ1' + str(count)] = lis[46]
            excle['BK1' + str(count)] = lis[47]
            excle['BL1' + str(count)] = lis[48]
            excle['BM1' + str(count)] = lis[49]
            excle['BN1' + str(count)] = lis[50]
            excle['BO1' + str(count)] = lis[51]
            excle['BP1' + str(count)] = lis[52]
            excle['BQ1' + str(count)] = lis[53]
            excle['BR1' + str(count)] = lis[54]
            excle['BS1' + str(count)] = lis[55]
            excle['BT1' + str(count)] = lis[56]
            excle['BU1' + str(count)] = lis[57]
            excle['BV1' + str(count)] = lis[58]
            excle['BW1' + str(count)] = lis[59]
            excle['BX1' + str(count)] = lis[60]
            excle['BY1' + str(count)] = lis[61]
            excle['BZ1' + str(count)] = lis[62]
        elif len(lis) == 67:
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
            excle['R' + str(count)] = lis[23]
            excle['S' + str(count)] = lis[24]
            excle['T' + str(count)] = lis[25]
            excle['U' + str(count)] = lis[26]
            excle['V' + str(count)] = lis[27]
            excle['W' + str(count)] = lis[28]
            excle['X' + str(count)] = ''
            excle['Y' + str(count)] = ''
            excle['Z' + str(count)] = ''
            excle['AA' + str(count)] = ''
            excle['AB' + str(count)] = ''
            excle['AC' + str(count)] = ''
            excle['AD' + str(count)] = ''
            excle['AE' + str(count)] = ''
            excle['AF' + str(count)] = ''
            excle['AG' + str(count)] = ''
            excle['AH' + str(count)] = ''
            excle['AR' + str(count)] = ''
            excle['AJ' + str(count)] = ''
            excle['AK' + str(count)] = ''
            excle['AL' + str(count)] = ''
            excle['AM' + str(count)] = ''
            excle['AN' + str(count)] = ''
            excle['AO' + str(count)] = lis[30]
            excle['AP' + str(count)] = lis[29]
            excle['AQ' + str(count)] = lis[30]
            excle['AR' + str(count)] = lis[31]
            excle['AS' + str(count)] = lis[32]
            excle['AT' + str(count)] = lis[33]
            excle['AU' + str(count)] = lis[34]
            excle['AV' + str(count)] = lis[35]
            excle['AW' + str(count)] = lis[36]
            excle['AX' + str(count)] = lis[37]
            excle['AY' + str(count)] = lis[38]
            excle['AZ' + str(count)] = lis[39]
            excle['BA1' + str(count)] = lis[40]
            excle['BB1' + str(count)] = lis[41]
            excle['BC1' + str(count)] = lis[42]
            excle['BD1' + str(count)] = lis[43]
            excle['BE1' + str(count)] = lis[44]
            excle['BF1' + str(count)] = lis[45]
            excle['BG1' + str(count)] = lis[46]
            excle['BH1' + str(count)] = lis[47]
            excle['BI1' + str(count)] = lis[48]
            excle['BJ1' + str(count)] = lis[49]
            excle['BK1' + str(count)] = lis[50]
            excle['BL1' + str(count)] = lis[51]
            excle['BM1' + str(count)] = lis[52]
            excle['BN1' + str(count)] = lis[53]
            excle['BO1' + str(count)] = lis[54]
            excle['BP1' + str(count)] = lis[56]
            excle['BQ1' + str(count)] = lis[57]
            excle['BR1' + str(count)] = lis[58]
            excle['BS1' + str(count)] = lis[59]
            excle['BT1' + str(count)] = lis[60]
            excle['BU1' + str(count)] = lis[61]
            excle['BV1' + str(count)] = lis[62]
            excle['BW1' + str(count)] = lis[63]
            excle['BX1' + str(count)] = lis[64]
            excle['BY1' + str(count)] = lis[65]
            excle['BZ1' + str(count)] = lis[66]


def store2(result):
    global excle2, count2
    for lis in result:
        excle2['A' + str(count2)] = lis[0]
        excle2['B' + str(count2)] = lis[1]
        excle2['C' + str(count2)] = lis[2]
        excle2['D' + str(count2)] = lis[3]
        excle2['E' + str(count2)] = lis[4]
        excle2['F' + str(count2)] = lis[5]
        excle2['G' + str(count2)] = lis[6]
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
    temp = 0
    number_lis = ['大名城', '天津松江', '云南城投', '嘉华国际', '华南城', '太古地产', '远洋集团']
    for i in stock_result:
        if i[0] not in number_lis:
            continue
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
    wb.save('信息集合3.xlsx')
    wb2.save('销售集合3.xlsx')