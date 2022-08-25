import requests
from bs4 import BeautifulSoup

from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.title = "NadoSheet"

for page in range(1,100, 10):
    #print(page)
    raw = requests.get(r"https://search.naver.com/search.naver?where=news&sm=tab_pge&query=%EC%A3%BC%EC%8B%9D&sort=0&photo=0&field=0&pd=0&ds=&de=&cluster_rank=98&mynews=0&office_type=0&office_section_code=0&news_office_checked=&nso=so:r,p:all,a:all&start="+str(page))
    html = BeautifulSoup(raw.text, "html.parser")

    clips = html.select("ul.list_news > li")
    # clips = html.select("a.news_tit")
    #print(clips[0].text)

    # c = clips[0].select_one('a.news_tit')
    a = 0
    a = a + page
    for c in clips:
        a = a + 1
        title = c.select_one('a.news_tit')
        print(title.text)
        ws["A" + str(a)] = title.text
    
wb.save(r"C:\file_python\samp.xlsx")

# print(c.text)


# print(clips[0])
# 컨테이너 소스코드는 주석처리해줍니다.

# for cl in clips:
#     title = cl.select_one("dt.title")
#     title1 = cl.select_one("dd.chn")
#     title2 = cl.select_one("span.hit")
#     title3 = cl.select_one("span.like")
#     print(title.text.strip())
#     print(title1.text.strip())
#     print(title2.text.strip())
#     print(title3.text.strip())
# #print(soup.title)
# print(soup.title.get_text())
# print(soup.a.get_text())
# print(soup.a.attrs)
# print(soup.a["href"])

#sp_nws1 > div.news_wrap.api_ani_send > div > a
# print(soup.find("a",attrs={"class":"Nbtn_upload"}))
# print(soup.find(attrs={"class":"Nbtn_upload"}))

# print(soup.find("li",attrs={"class":"rank01"}))
# rank1 =soup.find("li",attrs={"class":"rank01"})
# rank2 =rank1.next_sibling.next_sibling
# rank3 =rank2.next_sibling.next_sibling

# print(rank1.a)

# print(rank3.a.get_text())
# # print(rank3.next_sibling)
# rank2 = rank3.previous_sibling.previous_sibling
# print(rank2.next_sibling)
# print(rank1)