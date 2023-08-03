from selenium import webdriver
from selenium.webdriver.common.by import By
#from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import openpyxl
import requests
import json

#X-NCP-APIGW-API-KEY-ID': 'ku93d0gqaq, X-NCP-APIGW-API-KEY': 'vsdzhy0I0npUHmTn8mj1cKSKYVudzOryL7TYN9s7

i = 1
real_list = []
list_data = []
Area_data = []

search_business = input('검색할 사업을 입력하세요 : ')
curr_date = datetime.strftime(datetime.now(), "%y_%m_%d")
#browser = webdriver.Chrome(ChromeDriverManager().install())

# 부산지역 주요 지하철역 위도, 경도
Area_filename = '/Volumes/ext1/차진옥/python/pusan_subway_st1.xlsx'
Area_book = openpyxl.load_workbook(Area_filename)
Area_sheet = Area_book.worksheets[0]
for row in Area_sheet.rows:
    Area_data.append([
        row[1].value,
        row[2].value,
        row[3].value
                 ])
# 지하철명, 위도, 경도
for num in range(len(Area_data)-1):
    st_name = Area_data[num+1][0]
    lati = Area_data[num +1][1]
    logi = Area_data[num+1][2]

    for i in range(3):
      url = 'https://map.naver.com/v5/api/search?query=' + search_business + '&type=place&searchCoord=' + str(logi) + ';' + str(lati) + '&page=' + str(i+1) + '&displayCount=50'
      payload = json.dumps({
          "location": "https://naveropenapi.apigw.ntruss.com/map-static/v2/raster/ncpclientid=ku93d0gqaq",
          # "X-NCP-APIGW-API-KEY-ID": "ku93d0gqaq",
          "X-NCP-APIGW-API-KEY": "vsdzhy0I0npUHmTn8mj1cKSKYVudzOryL7TYN9s7",
          "argument": {
              "type": "1"
          }
      })
      headers = {
        'authority': 'map.naver.com',
        'method': 'GET',
        'scheme': 'https',
        'accept': 'application/json, text/plain, */*',
        'accept-encoding': 'gzip, deflate, br',
        'accept-language': 'ko-KR,ko;q=0.8,en-US;q=0.6,en;q=0.4',
        'cache-control': 'no-cache',
        }
      response = requests.get(url, headers=headers, data=payload)
      text = response.text
      data = json.loads(text)
      for list in data['result']['place']['list']:
        name = list['name']
        address = list['address']
        id = list['id']
        category = list['category'][0]
        if [name, address, id, category] in list_data:
            continue
        else:
            list_data.append([st_name, name, address, id, category])
columns = ['지하철역', '카페이름', '주소', 'ID', '구분']
real_df = pd.DataFrame(list_data, columns=columns)
real_df.head()
real_df.info()
real_df.to_excel('./부산지하철 주변지역 ' + search_business + ' ID ' + curr_date + '.xlsx', index=True)