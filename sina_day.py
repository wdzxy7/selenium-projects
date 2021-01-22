from bs4 import BeautifulSoup
from openpyxl import Workbook
import jieba
import re


def crawl():
    global words
    soup = BeautifulSoup(str(html), 'html.parser')
    divs = soup.find_all('div', {'class': 'WB_text W_f14'})
    for div in divs:
        mess = deal(str(div))
        mess = mess.replace('\\u200b', '').replace('\\ue627', '').replace('打工人超话', '').replace(' ', '').replace('\n', '')
        if len(mess) == 0:
            continue
        words.append(mess)


def deal(s):
    pattern = re.compile(r'<[^>]*>')
    remove = re.findall(pattern, s)
    for j in remove:
        s = s.replace(j, '')
    return s.replace('\n', '')


if __name__ == '__main__':
    work_dict = {
        '劳务': 0,
        '合同': 0,
        '零工': 0,
        '薪资': 0,
        '客户': 0,
        '开除': 0,
        '加班': 0,
        '上班': 0,
        '发票': 0,
        '迟到': 0,
        '考勤': 0,
        '挣钱': 0,
        '工地': 0,
        '社畜': 0,
        '简历': 0,
        '辞职': 0,
        '待遇': 0,
        '福利': 0,
        '下班': 0,
        '准时': 0,
        '领导': 0,
        '假': 0
    }
    study_dict = {
        '考研': 0,
        '毕业': 0,
        '复习': 0,
        '调研': 0,
        '论文': 0,
        '答辩': 0,
        '导师': 0,
        '同学': 0,
        '作业': 0,
        '读书': 0,
    }
    health_dict = {
        '干饭': 0,
        '胖': 0,
        '减肥': 0,
        '感冒': 0,
        '生病': 0,
        '颈椎': 0,
        '中药': 0,
        '嗓子': 0,
        '吃': 0,
        '头发': 0,
        '晚餐': 0,
        '早餐': 0,
        '睡': 0,
        '熬夜': 0,
        '瘦': 0,
    }
    transport_dict = {
        '地铁': 0,
        '打车': 0,
        '公交': 0,
        '买房': 0,
        '租房': 0,
        '房租': 0,
        '房东': 0,
    }
    bad_dict = {
        '惨': 0,
        '悲': 0,
        '焦虑': 0,
        '累': 0,
        '难过': 0,
        '可怜': 0,
        '烦': 0,
        '死': 0,
        '难': 0,
        '不值得': 0,
        '安慰': 0,
        '辛苦': 0,
        '奔波': 0,
        '压力': 0,
    }
    good_dict = {
        '勤恳': 0,
        '努力': 0,
        '加油': 0,
        '相信': 0,
        '美好': 0,
        '愉快': 0,
        '快乐': 0,
        '早安': 0,
        '晚安': 0,
        '顺利': 0,
        '喜欢': 0,
        '人上人': 0,
        '元气': 0,
        '坚持': 0,
        '收获': 0
    }
    else_dict = {
        '景色': 0,
        '老乡': 0,
        '周一': 0,
        '周五': 0,
        '周末': 0,
        '疫情': 0,
        '外卖': 0,
        '朋友': 0,
        '冬天': 0,
        '逃离': 0,
        '再见': 0,
        '一线': 0,
        '四线': 0,
        '假期': 0,
        '年假': 0,
    }
    words = []
    front = './新建文件夹/新建文本文档 ('
    back = ').txt'
    for i in range(1, 23):
        if i == 1:
            path = './新建文件夹/新建文本文档.txt'
        else:
            path = front + str(i) + back
        with open(path, 'r', encoding='utf-8') as f:
            html = f.readlines()
        crawl()


    wb1 = Workbook()
    excle1 = wb1.active
    excle1['A1'] = 'word'
    excle1['B1'] = 'count'
    count = 2
    for word in words:
        excle1['A' + str(count)] = word
        count = count + 1
    wb1.save('./新建文件夹/message.xlsx')
    wb1.close()