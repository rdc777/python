from bs4 import BeautifulSoup
import requests
import re
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import selenium
import pyautogui
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import datetime

from openpyxl import Workbook
browser = webdriver.Chrome(r"C:\file_python\chromedriver")
browser.implicitly_wait(10)
wb = Workbook()
ws = wb.active
ws.title = "웹 크롤링"

def makePgNum(num):
    if num == 1:
        return num
    elif num == 0:
        return num+1
    else:
        return num+9*(num-1)

def makeUrl(search,start_pg,end_pg):
    if start_pg == end_pg:
        start_page = makePgNum(start_pg)
        url = "https://search.naver.com/search.naver?where=news&sm=tab_pge&query=" + search + "&start=" + str(start_page)
        print("생성url: ",url)
        return url
    else:
        urls= []
        for i in range(start_pg,end_pg+1):
            page = makePgNum(i)
            url = "https://search.naver.com/search.naver?where=news&sm=tab_pge&query=" + search + "&start=" + str(page)
            urls.append(url)
        print("생성url: ",urls)
        return urls

search = input("검색할 키워드를 입력해주세요:")

#검색 시작할 페이지 입력
page = int(input("\n크롤링할 시작 페이지를 입력해주세요. ex)1(숫자만입력):")) # ex)1 =1페이지,2=2페이지...
print("\n크롤링할 시작 페이지: ",page,"페이지")   
#검색 종료할 페이지 입력
page2 = int(input("\n크롤링할 종료 페이지를 입력해주세요. ex)1(숫자만입력):")) # ex)1 =1페이지,2=2페이지...
print("\n크롤링할 종료 페이지: ",page2,"페이지")   

# naver url 생성
search_urls = makeUrl(search,page,page2)
b = 0
for page in search_urls:
    #print(page)
    raw = requests.get(page)
    html = BeautifulSoup(raw.text, "html.parser")



    clips = html.select("ul.list_news > li")##sp_nws99//*[@id="sp_nws99"]#sp_nws99
    # clips = html.select("a.news_tit")
    #print(clips[0].text)

    # c = clips[0].select_one('a.news_tit')
    b = b+1
    a = 0
    a = a + b
    for c in clips:
        a = a + 1
        title = c.select_one('a.news_tit')
        print(title.text)
        ws["A" + str(a)] = title.text
    
wb.save(r"C:\file_python\samp.xlsx")
 # 브라우저 종료하기
browser.quit()
