import time
from selenium import webdriver
import xlrd
import xlutils.copy
count = 1
cols = 0
tag = 0


class Getout(Exception):
    pass


def log(driver):
    xpath = '//div[@style=\'width:56px;float:left;text-align:center;padding-top:30px;\']'
    driver.find_element_by_xpath(xpath).click()
    xpath = '//input[@placeholder=\'账号\']'
    driver.find_element_by_xpath(xpath).send_keys('liuning')
    xpath = '//input[@placeholder=\'密码\']'
    driver.find_element_by_xpath(xpath).send_keys('Ln632080907')
    time.sleep(3)
    xpath = '//button[@class=\'signin\']'
    driver.find_element_by_xpath(xpath).click()


def go_to_need_page(driver):
    time.sleep(2)
    try:
        driver.switch_to.alert.accept()
    except:
        pass
    xpath = '//span[text()=\'销售管理系统\']'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(1)
    xpath = '//span[text()=\'客户检索\']'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(1)
    xpath = '//span[text()=\'添加新客户(新)\']'
    driver.find_element_by_xpath(xpath).click()
    time.sleep(10)


def send_message(driver, message, telephone):
    global tag
    print(message, telephone)
    time.sleep(2)
    driver.switch_to.default_content()
    iframe = driver.find_element_by_id('centerPage')
    driver.switch_to.frame(iframe)
    xpath = '//input[@name=\'TELEPHONE\']'
    input = driver.find_element_by_xpath(xpath)
    input.clear()
    xpath = '//input[@name=\'ACCOUNT_NAME\']'
    input = driver.find_element_by_xpath(xpath)
    input.clear()
    input.send_keys(message)
    xpath = '//input[@name=\'searchBtn\']'
    driver.find_element_by_xpath(xpath).click()
    try:
        xpath = '//td[@class=\'lisTdPROPERTY\']'
        results = driver.find_elements_by_xpath(xpath)
        open_list = []
        for result in results:
            open = result.get_attribute('textContent')
            open_list.append(open)
        xpath = '//td[@class=\'lisTdACCOUNT_LEVEL\']'
        results = driver.find_elements_by_xpath(xpath)
        status_list = []
        for result in results:
            status = result.get_attribute('textContent')
            status_list.append(status)
        result_list = []
        for i in range(len(open_list)):
            tup = (open_list[i], status_list[i])
            result_list.append(tup)
        if len(result_list) == 0:
            raise Getout()
        return result_list
    except Getout:
        tag = 1
        xpath = '//input[@name=\'ACCOUNT_NAME\']'
        input = driver.find_element_by_xpath(xpath)
        input.clear()
        xpath = '//input[@name=\'TELEPHONE\']'
        input = driver.find_element_by_xpath(xpath)
        input.clear()
        input.send_keys(telephone)
        xpath = '//input[@name=\'searchBtn\']'
        driver.find_element_by_xpath(xpath).click()
        try:
            xpath = '//td[@class=\'lisTdPROPERTY\']'
            results = driver.find_elements_by_xpath(xpath)
            open_list = []
            for result in results:
                open = result.get_attribute('textContent')
                open_list.append(open)
            xpath = '//td[@class=\'lisTdACCOUNT_LEVEL\']'
            results = driver.find_elements_by_xpath(xpath)
            status_list = []
            for result in results:
                status = result.get_attribute('textContent')
                status_list.append(status)
            result_list = []
            for i in range(len(open_list)):
                tup = (open_list[i], status_list[i])
                result_list.append(tup)
            return result_list
        except:
            return None


def get_name(sheet):
    result = []
    tels = []
    message_sum = sheet.nrows + 1
    for i in range(1, message_sum):
        name = sheet.cell_value(i-1, 0)
        telephone = sheet.cell_value(i-1, 1)
        if name == '' or name == '名称':
            continue
        result.append(name)
        tels.append(telephone)

    return result, tels


def store(table, message):
    global count, cols
    table.write(count, cols - 1, message)
    count = count + 1


def next(driver, telephone):
    driver.switch_to.default_content()
    iframe = driver.find_element_by_id('centerPage')
    driver.switch_to.frame(iframe)
    xpath = '//input[@name=\'ACCOUNT_NAME\']'
    input = driver.find_element_by_xpath(xpath)
    input.clear()

    xpath = '//input[@name=\'TELEPHONE\']'
    input = driver.find_element_by_xpath(xpath)
    input.clear()
    input.send_keys(telephone)

    xpath = '//input[@name=\'searchBtn\']'
    driver.find_element_by_xpath(xpath).click()
    try:
        xpath = '//td[@class=\'lisTdPROPERTY\']'
        results = driver.find_elements_by_xpath(xpath)
        open_list = []
        for result in results:
            open = result.get_attribute('textContent')
            open_list.append(open)
        xpath = '//td[@class=\'lisTdACCOUNT_LEVEL\']'
        results = driver.find_elements_by_xpath(xpath)
        status_list = []
        for result in results:
            status = result.get_attribute('textContent')
            status_list.append(status)
        result_list = []
        for i in range(len(open_list)):
            tup = (open_list[i], status_list[i])
            result_list.append(tup)
        return result_list
    except:
        return None


if __name__ == '__main__':
    front = 'H://python//'
    back = '111.xls'
    path = front + back
    data = xlrd.open_workbook(path)
    sheet = data.sheet_by_name("Sheet")
    cols = sheet.ncols + 1
    print(cols)
    result, telephones = get_name(sheet)
    ws = xlutils.copy.copy(data)
    table = ws.get_sheet(0)
    url = 'http://sales.vemic.com/'
    driver = webdriver.Chrome()
    driver.get(url)
    log(driver)
    go_to_need_page(driver)
    for i in range(len(result)):
        pdd = 0
        message = str(result[i])
        telephone = str(telephones[i]).replace('.0', '')
        tag = 0
        res = send_message(driver, message, telephone)
        print(res)
        if res is None or len(res) == 0:
            lable = '1'
            store(table, lable)
            continue
        if tag == 1:
            tup = res[0]
            open_mes = tup[0]
            if open_mes == '开放':
                lable = '2'
            else:
                lable = '3'
            store(table, lable)
            print(message, open_mes,  lable)
            continue
        lable = '2'
        if len(res) == 1:
            tup = res[0]
            open_mes = tup[0]
            level = tup[1]
            if open_mes == '开放':
                ress = next(driver, telephone)
                if ress is None or len(ress) == 0:
                    lable = '2'
                else:
                    tup = ress[0]
                    op = tup[0]
                    if op == '开放':
                        lable = '2'
                    else:
                        lable = '3'
                print(message, open_mes, level, lable)
                store(table, lable)
                continue
            else:
                if level == '一般潜在' or level == '放弃' or level == '未联系' or level == '系统开放' or level == '空':
                    lable = '4'
                elif level == '流失' or level == '过期流失':
                    lable = '4'
                elif level == '优质潜在' or level == '待签':
                    lable = '5'
                elif level == '签约':
                    lable = '6'
        else:
            for i in range(len(res)):
                tup = res[i]
                open_mes = tup[0]
                level = tup[1]
                tlable = '0'
                if open_mes == '开放':
                    ress = next(driver, telephone)
                    if ress is None or len(ress) == 0:
                        lable = '2'
                    else:
                        tup = ress[0]
                        op = tup[0]
                        if op == '开放':
                            lable = '2'
                        else:
                            lable = '3'
                    print(message, open_mes, level, lable)
                    store(table, lable)
                    pdd = 1
                    break
                else:
                    if level == '一般潜在' or level == '放弃' or level == '未联系' or level == '系统开放' or level == '空':
                        tlable = '4'
                    elif level == '流失' or level == '过期流失':
                        tlable = '4'
                    elif level == '优质潜在' or level == '待签':
                        tlable = '5'
                    elif level == '签约':
                        tlable = '6'
                    if int(tlable) > int(lable):
                        lable = tlable
        time.sleep(1)
        if pdd != 1:
            store(table, lable)
            print(message, open_mes, level, lable)
    ws.save('H://桌面文件//' + back)
    driver.quit()
