from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import selenium.webdriver.support.ui as ui
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
import requests
import threading
import re
import time


#主函数
def main():
    print('测试版本，由于没有使用账号密码登录')
    print('使用前必须使用chrome登陆过qq空间！！！')
    qq_number = input('输入qq号:')
    ################ 以下两句话可以适当修改
    chrome_options = Options()
    driver = webdriver.Chrome(executable_path='D:\RJY\别的\生产实习\chromedriver.exe', options=chrome_options)
    #################
    url = f'https://user.qzone.qq.com/{qq_number}/311'
    driver.get(url)
    driver.switch_to.frame('login_frame') # 发现之后进入login_frame
    time.sleep(1) # 响应时间

    login = driver.find_element_by_id(f'img_out_{qq_number}') # 获取点击按钮  也可以进行输入账号密码
    login.click() # 进行点击
    time.sleep(1) # 响应时间

    button = driver.find_element_by_id('aIcenter')
    button.click() # 进行点击
    time.sleep(1)

    driver.get(url)
    wait = ui.WebDriverWait(driver, 10)
    wait.until(lambda driver: driver.find_element_by_xpath('//*[@id="menuContainer"]/div/ul/li[5]'))
    button = driver.find_element_by_xpath('//*[@id="menuContainer"]/div/ul/li[5]')  # 获取说说按钮  也可以进行输入账号密码
    button.click() # 进行点击
    time.sleep(1)

    driver.switch_to.default_content()
    frame = driver.find_elements_by_tag_name('iframe')[0]
    driver.switch_to.frame(frame)
    find_msg_List = driver.find_elements_by_xpath('//*[@id="msgList"]/li')
    msg_ListNum = len(find_msg_List)
    print(msg_ListNum)

    for i in range(1,msg_ListNum+1):
        while 1:
            time.sleep(0.1)
            try:
                find_msg_List = driver.find_element_by_xpath('//*[@id="msgList"]/li[%d]/div[3]/div[2]'%i)
                print('已定位到元素%s'%i)
                break
            except:
                print("还未定位到元素%s!"%i)
        find_msg_text = find_msg_List.text
        print(find_msg_text)
    #find_msgList = driver.find_element_by_xpath('//*[@id="msgList"]/li[1]/div[3]/div[2]')

    driver.quit()
    print('WellDone')
    return



#执行函数
if __name__ == '__main__':
    main()
