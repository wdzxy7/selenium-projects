from selenium import webdriver
import random


def multiple_choice(driver, id, question_num):
    frequency = random.randint(1, 5)
    choiced_list = set()
    for i in range(frequency):
        choice_num = random.randint(1, 5)
        if choice_num in choiced_list:
            continue
        else:
            choiced_list.add(choice_num)
        question = '\'q' + str(question_num) + '_' + str(choice_num) + '\''
        xpath = '//div[@id=' + id + ']/div[@class=\'ui-controlgroup\']' \
                '/div[@class=\'ui-checkbox\']//input[@id=' + question + ']' \
                '/following-sibling::a[1]'
        try:
            driver.find_element_by_xpath(xpath).click()
        except:
            choice_num = random.randint(1, 3)
            if choice_num in choiced_list:
                continue
            else:
                choiced_list.add(choice_num)
            question = '\'q' + str(question_num) + '_' + str(choice_num) + '\''
            xpath = '//div[@id=' + id + ']/div[@class=\'ui-controlgroup\']' \
                    '/div[@class=\'ui-checkbox\']//input[@id=' + question + ']' \
                    '/following-sibling::a[1]'
            try:
                driver.find_element_by_xpath(xpath).click()
            except:
                choice_num = random.randint(1, 2)
                if choice_num in choiced_list:
                    continue
                else:
                    choiced_list.add(choice_num)
                question = '\'q' + str(question_num) + '_' + str(choice_num) + '\''
                xpath = '//div[@id=' + id + ']/div[@class=\'ui-controlgroup\']' \
                        '/div[@class=\'ui-checkbox\']//input[@id=' + question + ']' \
                        '/following-sibling::a[1]'
                driver.find_element_by_xpath(xpath).click()


def fill_questionnaire():
    driver = webdriver.Chrome()
    driver.get('https://www.wjx.cn/m/74452649.aspx')
    for question_num in range(1, 21):
        id = '\'div' + str(question_num) + '\''
        if question_num == 8 or question_num > 14:
            multiple_choice(driver, id, question_num)
            continue
        choice_num = random.randint(1, 5)
        question = '\'q' + str(question_num) + '_' + str(choice_num) + '\''
        xpath = '//div[@id=' + id + ']/div[@class=\'ui-controlgroup\']' \
                '/div[@class=\'ui-radio\']//input[@id=' + question + ']' \
                '/following-sibling::a[1]'
        try:
            driver.find_element_by_xpath(xpath).click()
        except:
            choice_num = random.randint(1, 3)
            question = '\'q' + str(question_num) + '_' + str(choice_num) + '\''
            xpath = '//div[@id=' + id + ']/div[@class=\'ui-controlgroup\']' \
                    '/div[@class=\'ui-radio\']//input[@id=' + question + ']' \
                    '/following-sibling::a[1]'
            try:
                driver.find_element_by_xpath(xpath).click()
            except:
                choice_num = random.randint(1, 2)
                question = '\'q' + str(question_num) + '_' + str(choice_num) + '\''
                xpath = '//div[@id=' + id + ']/div[@class=\'ui-controlgroup\']' \
                        '/div[@class=\'ui-radio\']//input[@id=' + question + ']' \
                        '/following-sibling::a[1]'
                driver.find_element_by_xpath(xpath).click()
    submit = '//a[@id=\'ctlNext\']'
    driver.find_element_by_xpath(submit).click()
    driver.quit()


if __name__ == '__main__':
    # 输入填写次数
    count = input('输入填写次数:')
    count = int(count)
    if count <= 0:
        print('执行次数小于0请输入争取的执行次数')
        count = input()
        count = int(count)
        if count <= 0:
            print('输入错误程序结束执行')
            count = 0
    for i in range(count):
        print('当前填写第' + str(i + 1) + '次')
        fill_questionnaire()