import selenium
import pyautogui
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import datetime

#options = webdriver.ChromeOptions()
#options.add_experimental_option("excludeSwitches", ["enable-logging"])
#browser = webdriver.Chrome(options=options)

browser = webdriver.Chrome(r"C:\file_python\chromedriver")
browser.implicitly_wait(10)
url = "https://www.hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index.xml"
#url = "https://www.google.com/?gws_rd=ssl"
browser.get(url)
browser.maximize_window()

id = browser.find_element_by_css_selector("#textbox81212912")
id.click()

browser.switch_to.frame('txppIframe')
time.sleep(1)
browser.find_element_by_xpath('//*[@id="anchor22"]').click()

#공인인증서 로그인
time.sleep(3)
browser.switch_to.frame('dscert')
time.sleep(1)
ex = browser.find_element_by_css_selector('#input_cert_pw')
ex.click()

import tkinter
from tkinter import messagebox

#def btnClick():
#    global txt
#    txt = text.get()
#    win.destroy()
    
    

#win = tkinter.Tk()
#win.title("공인인증서 비밀번호 입력")
#win.geometry('300x90')
#text = tkinter.Entry(win)
#text.grid(row=1, column=10, padx=65)
#button = tkinter.Button(win, text="확인", command=btnClick)
#button.grid(row=4, column=10, padx=65)

#win.mainloop()

#ex.send_keys(txt)
ex.send_keys(
'wjstks1713!'
)
browser.find_element_by_css_selector("#btn_confirm_iframe > span").click()

time.sleep(1)

#조회/발급
browser.find_element_by_xpath('//*[@id="textbox81212923"]').click()
time.sleep(1)
browser.switch_to.frame('txppIframe')
time.sleep(1)
browser.find_element_by_xpath('//*[@id="sub_a_0108020000"]').click()

#약간동의
time.sleep(1)
browser.find_element_by_css_selector('#radioYn_input_0').click()
browser.find_element_by_css_selector('#trigger7').click()

#사업자등록조회

time.sleep(2)

#browser.switch_to.parent('txppIframe')

browser.find_element_by_xpath('//*[@id="radio_input_0"]').click()

nw = browser.find_element_by_xpath('//*[@id="agrDt_input"]')
nw.click()

from datetime import datetime

a=datetime.today()
y=str(a.year)
m=str(a.month).zfill(2)
d=str(a.day).zfill(2)

nw.send_keys(y+m+d)


ta = browser.find_element_by_xpath('//*[@id="inqrPrpse"]')
ta.click()
ta.send_keys("")#주민등록번호로 사업자등록하는 사유

from openpyxl import load_workbook
import os

os.chdir("")#반영할 엑셀 주소
wb = load_workbook(filename = '엑셀온라인 신청 확인.xlsm') # 파일 연동
sr = wb['온라인신청명단'] # 워크시트 호출

cel = 4
row_b = sr["P1"].value

while cel <= row_b:
    i = str(cel)
    ac = sr['N'+i].value
    ab = sr['O'+i].value
   
    
    wn1 = browser.find_element_by_id('resno1')
    wn1.send_keys(ac)

    wn2 = browser.find_element_by_id('resno2')
    
    wn2.send_keys(ab)

    browser.find_element_by_xpath('//*[@id="trigger5"]').click()
    #내용확인
    time.sleep(1)
    val = browser.find_element_by_xpath('//*[@id="grid2_cell_0_1"]/span')
    val2 = val.text
    ab = sr['O'+i].value
    sr["P"+i] = val2
    print(ac,ab)
    print(val2)
    print(row_b-3,"/",cel-3)
    time.sleep(2)
    cel = cel + 1
    wn1.clear()
    wn2.clear()
    
del wb['다운로드']
del wb['2019']
del wb['2020']
del wb['2021']
del wb['2021(추경)']
del wb['2022']    
wb.save('') # 엑셀 파일 저장

#


