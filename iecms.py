from bs4 import BeautifulSoup
import requests
import xlrd


def get_name(sheet):
    result = []
    message_sum = sheet.nrows + 1
    for i in range(2, message_sum):
        name = sheet.cell_value(i-1, 0)
        result.append(name)
    return result


def crawl(html):
    soup = BeautifulSoup(html, 'html.parser')
    print(soup)
    a = soup.find('a', {'id': 'ExternalLink'})
    print(a)


async def get_page():
    pass


if __name__ == '__main__':
    path = 'company.xls'
    data = xlrd.open_workbook(path)
    sheet = data.sheet_by_name("Sheet")
    search_list = get_name(sheet)
    for name in search_list:
        url = 'https://iecms.mofcom.gov.cn/pages/corp/NewCMvCorpInfoTabList.html?sp=S&sp=S' + name + '&sp=S&sp=Ssearch&sp=S8308'
        print(name, url)
        break