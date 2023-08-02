from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
from selenium.webdriver.common.keys import Keys
import openpyxl
# 네이버 부동산 실행
url = 'https://land.naver.com/'
curr_date = datetime.strftime(datetime.now(), "%y_%m_%d")
browser = webdriver.Chrome(ChromeDriverManager().install())
browser.get(url)
url = input('검색지역의 url을 복사해 넣으세요 : ')
browser.get(url)
# 조회할 지역의 자료를 페이지다운으로 스크롤 하고 지역명 입력
search_area = input("검색지역명을 입력하세요 : ")
time.sleep(2)

real_list = []
i = 0
body = browser.find_element_by_css_selector('#listContents1 > div > div > div:nth-child(1) > div:nth-child(1) > div > a')
for i in range(30):
    body.send_keys(Keys.PAGE_DOWN)

html = browser.page_source
soup = BeautifulSoup(html, 'html.parser')
lists = soup.find_all("div", {"class": "item"})
for list in lists:
    title = list.find("div", {"class": "item_title"})
    if title == None:
        title = 0
    else:
        title = title.text.strip()
    price = list.find('span', {"class": "price"})
    if price.text.find('억') > 0:
        sale_price_list = price.text.replace(',', '').split('억')
        if sale_price_list[1] == '':
            sale_price = int(sale_price_list[0]) * 10000
        else:
            sale_price = int(sale_price_list[0]) * 10000 + int(sale_price_list[1])
    else:
        sale_price = int(price.text.replace(',', ''))
    specs_list = list.find('span', {"class": "spec"})
    area = specs_list.text.replace('m²', '')
    regist = list.find('em', {"class": "data"})
    if regist == None:
        regist = "-"
    else:
        regist = regist.text.strip()
    area = round(float(area) / 3.3, 1)
# 전용면적 기준으로 평당 가격을 산출
    p_sale_price = round(sale_price / area, 0)
    real_list.append([title, sale_price, area, p_sale_price, regist])
columns = ['구분', '매매가', '계약면적', '평당매매가', '등록일']
real_df = pd.DataFrame(real_list, columns=columns)
real_df.head()
real_df.info()
real_df.to_excel('./'+ search_area +'주변 토지매물가_list ' + curr_date + '.xlsx', index=True)