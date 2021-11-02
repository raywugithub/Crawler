# coding: utf-8
"""
Post the query to Google　Search and get the return results
"""
import re
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pandas as pd

# Browser settings
chrome_options = Options()
chrome_options.add_argument('--incognito')
chrome_options.add_argument(
    'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36')
# 今天講個特別的，我們可以不讓瀏覽器執行在前景，而是在背景執行（不讓我們肉眼看得見）
# 如以下宣告 options
chrome_options.add_argument('--headless')

chrome_options.add_experimental_option(
    "excludeSwitches", ['enable-automation', 'enable-loggin'])

browser = webdriver.Chrome(
    chrome_options=chrome_options)


# browser.get(
#    'http://moneydj.jihsun.com.tw/z/zk/zkf/zkResult.asp?A=x@270,a@4000;x@1410,a@5,b@10;x@370,a@1,b@400&B=x@17,SO@0&D=1&site=')
#
#
# Crawler
#soup = BeautifulSoup(browser.page_source, 'html.parser')
#
# print(soup)

table = pd.read_html(
    'http://moneydj.jihsun.com.tw/z/zk/zkf/zkResult.asp?A=x@270,a@3000;x@1410,a@5,b@10;x@370,a@1,b@300&B=x@17,SO@0&D=1&site=')

#table = pd.read_html('https://www.taifex.com.tw/cht/2/stockLists')

print(table)

# Close the browser
browser.close()
