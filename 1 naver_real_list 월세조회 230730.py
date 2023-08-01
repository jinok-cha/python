from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import openpyxl
# 네이버 부동산 실행
url = 'https://land.naver.com/'
curr_date = datetime.strftime(datetime.now(), "%y_%m_%d")

option = webdriver.ChromeOptions()
option.add_argument("start-maximized")
browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=option)

browser.get(url)
url = input('검색지역의 url을 복사해 넣으세요 : ')
browser.get(url)
# 조회할 지역의 자료를 페이지다운으로 스크롤 하고 지역명 입력
search_area = input("물건을 조회하고 검색지역명을 입력하세요 : ")
time.sleep(2)

body = browser.find_element(By.CSS_SELECTOR, '#listContents1 > div > div > div:nth-child(1) > div:nth-child(1) > div > a')
for i in range(50):
    body.send_keys(Keys.PAGE_DOWN)
time.sleep(1)

real_list = []
i = 0
gubun = ''

html = browser.page_source
soup = BeautifulSoup(html, 'html.parser')
lists = soup.find_all("div", {"class": "item"})
for list in lists:
    print(list)
    title = list.find("div", {"class": "item_title"})
    if title == None:
        title = 0
    else:
        title = title.text.strip()
    price = list.find('span', {"class": "price"})
    if price == None:
        price = 0
    else:
        price = price.text.strip()
    price = str(price).split("/")
    bo_price = price[0]
    bo_price = bo_price.replace('억', '0000')
    bo_price = bo_price.replace('0000 ', '')
    if len(price) == 2:
        mon_price = price[1]
        gubun = '월세'
        sale_price = '0'
    elif len(price) == 1:
        sale_price = bo_price
        gubun = "매매"
        bo_price = '0'
        mon_price = '0'
    else:
        sale_price = '0'
        gubun = '월세'
        bo_price = '0'
        mon_price = '0'
    specs_list = list.find('span', {"class": "spec"})
    if specs_list == None:
        specs_list = "-"
    else:
        specs_list = specs_list.text.strip()
    specs = specs_list.split(',')
    if len(specs) > 1:
        spec = specs[0]
        chung = specs[1]
    area = spec.split('/')
    regist = list.find('em', {"class": "data"})
    if regist == None:
        regist = "-"
    else:
        regist = regist.text.strip()
    sale_price = int(sale_price.replace(',', ''))
    bo_price = int(bo_price.replace(',', ''))
    mon_price = int(mon_price.replace(',', ''))
    area[0] = round(float(area[0].replace(',', '')) / 3.3, 1)
    area[1] = round(float(area[1].replace('m²', '')) / 3.3, 1)
# 전용면적 기준으로 평당 가격을 산출
    p_sale_price = round(sale_price/area[1], 0)
    p_bo_price = round(bo_price/area[1], 0)
    p_mon_price = round(mon_price/area[1], 1)
    real_list.append([gubun, title, sale_price, bo_price, mon_price, area[0], area[1], chung, p_sale_price, p_bo_price, p_mon_price, regist])
columns = ['구분', '종류', '매매가', '보증금', '월세', '계약면적', '전용면적', '층수', '평당매매가', '평당보증금', '평당월세', '등록일']
real_df = pd.DataFrame(real_list, columns=columns)
real_df.head()
real_df.info()
real_df.to_excel('./'+ search_area +'월세조회_list ' + curr_date + '.xlsx', index=True)