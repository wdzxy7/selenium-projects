import re
import xlrd
from bs4 import BeautifulSoup
from openpyxl import Workbook
from requests.exceptions import RequestException
import requests

job_name_list = []
company_name_list = []
work_place_list = []
salary_list = []
urls = []
worked_time_list = []
education_list = []
industry_list = []


# 读取excle获得url
def get_url(sheet):
    result = []
    message_sum = sheet.nrows + 1
    for i in range(2, message_sum):
        name = sheet.cell_value(i-1, 0)
        result.append(name)
    return result


# 发送请求返回网页源码
def get_html(url):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.162 Safari/537.36'
        }
        response = requests.get(url, headers=headers)
        response.encoding = 'gbk'
        if response.status_code == 200:
            return response.text
        return None
    except RequestException:
        return None


# 爬取信息 用bs4 根据属性获取
def crawl(html):
    wb = Workbook()
    excle = wb.active
    creat_wb(excle)
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('div', {'class': 'dw_table'})
    soup = BeautifulSoup(str(table), 'html.parser')
    divs = soup.find_all('div', {'class': 'el'})
    del divs[0]
    for div in divs:
        soup = BeautifulSoup(str(div), 'html.parser')
        a = soup.find('a')
        job = a.contents[0]
        url = a.get('href')
        company_name = soup.find('span', {'class': 't2'}).contents[0]
        company_name = deal(str(company_name))
        print(company_name)
        try:
            salary = soup.find('span', {'class': 't4'}).contents[0]
        except:
            salary = 'None'
        job_name_list.append(str(job).strip())
        urls.append(str(url).strip())
        company_name_list.append(company_name)
        salary_list.append(str(salary).strip())


# 处理掉html标签
def deal(s):
    pattern = re.compile(r'<[^>]*>')
    # 匹配html标签
    remove = re.findall(pattern, s)
    # 处理
    for i in remove:
        s = s.replace(i, '')
    return s


# 进一步获取信息
def deep_crawl(html):
    if html is None:
        work_place_list.append('None')
        worked_time_list.append('None')
        education_list.append('None')
        industry_list.append('None')
        return None
    soup = BeautifulSoup(html, 'html.parser')
    p = soup.find('p', {'class': 'msg ltype'})
    print(p)
    if p is None:
        work_place_list.append('None')
        worked_time_list.append('None')
        education_list.append('None')
        industry_list.append('None')
        return None
    # 获取信息
    title = p.get('title')
    lis = str(title).replace('\xa0', '').split('|')
    place = lis[0]
    worked_time = lis[1]
    education = lis[2]
    work_place_list.append(place)
    worked_time_list.append(worked_time)
    education_list.append(education)
    ps = soup.find_all('p', {'class': 'at'})
    p = ps[2]
    title = p.get('title')
    industry_list.append(title)


# 创建存储的excle表格
def creat_wb(excle):
    excle['A1'] = 'jobName'
    excle['B1'] = 'company'
    excle['C1'] = 'salary'
    excle['D1'] = 'city'
    excle['E1'] = 'jobtype'
    excle['F1'] = 'eduLevel'
    excle['G1'] = 'workingExp'


if __name__ == '__main__':
    wb = Workbook()
    excle = wb.active
    creat_wb(excle)
    path = 'url_list.xlsx'
    data = xlrd.open_workbook(path)
    sheet = data.sheet_by_name("Sheet1")
    url_list = get_url(sheet)
    for url in url_list:
        html = get_html(url)
        crawl(html)
    for url in urls:
        html = get_html(url)
        deep_crawl(html)
    count = 2
    # 存储进exlce
    for i in range(len(job_name_list)):
        key = 'A' + str(count)
        excle[key] = job_name_list[i]
        key = 'B' + str(count)
        excle[key] = company_name_list[i]
        key = 'C' + str(count)
        excle[key] = salary_list[i]
        key = 'D' + str(count)
        excle[key] = work_place_list[i]
        key = 'E' + str(count)
        excle[key] = industry_list[i]
        key = 'F' + str(count)
        excle[key] = education_list[i]
        key = 'G' + str(count)
        excle[key] = worked_time_list[i]
        count = count + 1
    wb.save('51job.xlsx')