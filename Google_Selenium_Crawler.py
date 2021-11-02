# coding: utf-8
"""
Post the query to Google　Search and get the return results
"""
import re
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


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

sleep_time = 5

# Query settings
query = '原油'
query = '漲 OR 跌 OR 恆生 OR 港股'
query = '銅 AND 倫敦'
query = '漲 OR 跌 -台股'
query = '比特 AND ETF'
query = '(漲 OR 跌) OR (需求 OR 供應) -台股 -中油'

next_page_times = 10
printlink = 0
tbs_selection = 'h'  # h / d / w / m / y


if tbs_selection == 'h':
    tbs = '&tbs=qdr:h'  # past hour
elif tbs_selection == 'd':
    tbs = '&tbs=qdr:d'  # past day
elif tbs_selection == 'w':
    tbs = '&tbs=qdr:w'  # past week
elif tbs_selection == 'm':
    tbs = '&tbs=qdr:m'  # past month
elif tbs_selection == 'y':
    tbs = '&tbs=qdr:y'  # past year

# browser.get('https://www.google.com/search?q={}'.format(query) +
#            '&tbs=qdr:h')
browser.get('https://www.google.com/search?q={}'.format(query) +
            tbs)

# Crawler
for _page in range(next_page_times):
    soup = BeautifulSoup(browser.page_source, 'html.parser')
    content = soup.prettify()

    # Get titles and urls
    titles = re.findall('<h3 class="[\w\d]{6} [\w\d]{6}">\n\ +(.+)', content)
    # VwiC3b yXK7lf MUxGbd yDYNvb lyLwlc lEBKkf
    # mCBkyc JQe2Ld nDgy9d
    # subtitles = soup.find_all(
    #    'div', class_='VwiC3b yXK7lf MUxGbd yDYNvb lyLwlc lEBKkf')
    # urls = re.findall(
    #    '<div class="r">\ *\n\ *<a href="(.+)" onmousedown', soup.prettify())
    time_ago = re.findall(
        '<span class="MUxGbd wuQ4Ob WZ8Tjf">\n\ +(.+)', content)
    link = re.findall('<div class="yuRUbf">\n\ +(.+)', content)

    print('Page:', _page+1, '...........')
    # for n in range(min(len(titles), len(urls))):
    #    print(titles[n], urls[n])
    for n in range(len(titles)):
        if len(titles) == len(time_ago):
            print(titles[n] + '...' + time_ago[n])
            if printlink == 1:
                print('......................' + link[n].split()[2][6:-1])
        else:
            print(titles[n] + '...')
            if printlink == 1:
                print('............................................' +
                      link[n].split()[2][6:-1])

    #print('Page:', _page, end='\n\n')

    # Wait
    time.sleep(sleep_time)

    # Turn to the next page
    try:
        browser.find_element_by_link_text('下一頁').click()
    except:
        print('Search Early Stopping.')
        browser.close()
        exit()


# Close the browser
browser.close()
