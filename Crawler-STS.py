import urllib.request as request
from bs4.dammit import html_meta
import requests as req
from bs4 import BeautifulSoup

# Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36

url_time_entries = "https://sts.aaeon.com.tw:12080/time_entries"

url_login = "https://sts.aaeon.com.tw:12080/login"
login_data = {
    "username": "raywu",
    "password": "@WSX2wsx"
}

time_entries_data = {
    "time_entry[issue_id]": "15080",
    "time_entry[spent_on]": "2021-09-08",
    "time_entry[hours]": "1",
    "time_entry[comments]": "STS Management",
    "time_entry[activity_id]": "2"
}

url_paramters = {'': '', '': ''}
url_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36"
}

session = req.session()
session.post(url_login, headers=url_headers, data=login_data)

post_status = session.post(
    url_time_entries, headers=url_headers, data=time_entries_data)

print([post_status])
#response = session.get(url_time_entries, headers=url_headers)
#
#soup = BeautifulSoup(response.text, 'lxml')
#
#divs = soup.find('div', id='content')
#
# print(divs.h2)
