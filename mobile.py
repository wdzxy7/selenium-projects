import re

from bs4 import BeautifulSoup
from requests.exceptions import RequestException
import requests


def get_html(url):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.162 Safari/537.36'
        }
        response = requests.get(url, headers=headers)
        response.encoding = 'utf-8'
        if response.status_code == 200:
            return response.text
        return None
    except RequestException:
        return None


def deal(s):
    pattern = re.compile(r'<[^>]*>')
    remove = re.findall(pattern, s)
    for i in remove:
        s = s.replace(i, '')
    return s.replace('\n', '')


if __name__ == '__main__':
    url = 'https://bbs.pku.edu.cn/123/'
    html = get_html(url)
    soup = BeautifulSoup(html, 'html.parser')
    a = soup.find_all('a', {'target': '_blank'})
    for i in a:
        url = i['href']
        name = deal(str(i))
        if len(name) == 1:
            break
        print(name, url)