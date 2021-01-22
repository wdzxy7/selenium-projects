import time
from selenium import webdriver
import selenium
import msvcrt,sys

from selenium.webdriver import ActionChains


def log(driver, user_name, password):
    url = 'https://i.qq.com/'
    driver.get(url)
    iframe = driver.find_element_by_name('login_frame')  # 发现登陆的frame
    driver.switch_to.frame(iframe)
    driver.find_element_by_xpath('//a[text()=\'帐号密码登录\']').click()
    driver.find_element_by_xpath('//*[@id="u"]').clear()
    driver.find_element_by_xpath('//*[@id="u"]').send_keys(user_name)
    driver.find_element_by_xpath('//*[@id="p"]').send_keys(password)
    driver.find_element_by_xpath('//*[@id="login_button"]').click()
    time.sleep(2)
    distance = 180
    while True:
        try:
            iframe = driver.find_element_by_xpath('//iframe')
            driver.switch_to.frame(iframe)
            button = driver.find_element_by_id('tcaptcha_drag_button')
            time.sleep(1)
            action = ActionChains(driver)
            action.reset_actions()
            action.click_and_hold(button).perform()
            action.move_by_offset(distance, 0).perform()
            action.release().perform()
            break
        except:
            distance = distance - 3
            if distance < 170:
                distance = 180
    ur = driver.current_url
    if ur == url:
        print('密码错误请重新输入')
        return False
    else:
        return True


def pwd_input():
    chars = []
    while True:
        try:
            newChar = msvcrt.getch().decode(encoding="utf-8")
        except:
            return input("你很可能不是在cmd命令行下运行，密码输入将不能隐藏:")
        if newChar in '\r\n': # 如果是换行，则输入结束
             break
        elif newChar == '\b': # 如果是退格，则删除密码末尾一位并且删除一个星号
             if chars:
                 del chars[-1]
                 msvcrt.putch('\b'.encode(encoding='utf-8')) # 光标回退一格
                 msvcrt.putch( ' '.encode(encoding='utf-8')) # 输出一个空格覆盖原来的星号
                 msvcrt.putch('\b'.encode(encoding='utf-8')) # 光标回退一格准备接受新的输入
        else:
            chars.append(newChar)
            msvcrt.putch('*'.encode(encoding='utf-8')) # 显示为星号
    return (''.join(chars) )


if __name__ == '__main__':
    user_name = input('请输入帐号:')
    while True:
        try:
            t = int(user_name)
            break
        except:
            print('输入账号错误')
            user_name = input()
    print('请输入密码:', end='')
    password = pwd_input()
    # password = input()
    driver = webdriver.Chrome()
    while True:
        if log(driver, user_name, password):
            break
        else:
            password = pwd_input()
            # password = input()
    print(f"""
                  ======================
                              
                         登录成功！
                                       
                   ======================""")