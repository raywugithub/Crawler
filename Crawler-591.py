import urllib.request as req
import requests

# Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36

url = "https://sale.591.com.tw/home/search/list?type=2&shType=list&regionid=3&section=44,26,43,47&price=$_1200$&area=18$_$&houseage=$_30$&order=posttime_desc&timestamp=1631088443435"
url = "https://rent.591.com.tw/"

url_paramters = {'': '', '': ''}
url_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36"
    #    "": "gAwHZpSoBLPV2ve26bcQeGgnXu0V6UcfP8jehP57"
}

#request = req.Request(url, headers=url_headers)
# print(request)
#request = requests.get(url, headers=url_headers)
# print(request)
#
# with req.urlopen(request) as response:
#    data = response.read().decode("utf-8")

session = requests.session()
session_result = session.get(url)
print(session_result)
