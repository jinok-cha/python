# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from webdriver_manager.chrome import ChromeDriverManager
# import openpyxl
# from urllib import parse
# import time
import pandas as pd
from datetime import datetime
import requests
import json

j = 1
list_data = []
ID_data = []
real_list = []

curr_date = datetime.strftime(datetime.now(), "%y_%m_%d")
#구글맵에서 지도상에 원하는 위치에 마우스 오른쪽을 눌러 위도와 경도를 클릭하면 복사됨
url = input('검색지역의 위도와 경도를 복사해 넣으세요 : ')
url1 = url.split(',')
search_area = input("검색지역명을 입력하세요 : ")
search_business = input("검색할 사업종류를 입력하세요 : ")

#한가지 위도에 대해 5번의 검색을 하고 위도를 변경함(북쪽으로 0.03만큼 이동)
for i in range(5):
  if i % 3 == 0:
      url1[0] = float(url1[0]) + 0.03
      i = 0
  url = 'https://map.naver.com/v5/api/search?query=' + search_business + '&type=place&searchCoord=' + str(url1[1]) + ';' + str(url1[0]) + '&page=' + str(i+1) + '&displayCount=50'
  payload = json.dumps({
      "location": "https://naveropenapi.apigw.ntruss.com/map-static/v2/raster/ncpclientid=ku93d0gqaq",
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
#추출된 id를 이용하여 매장의 정보를 검색
  if 'result' in data:
      for list in data['result']['place']['list']:
        id = list['id']
        url = 'https://map.naver.com/v5/api/sites/summary/' + str(id) + '?lang=ko'
        payload = {}
        headers = {}
        response = requests.request("GET", url, headers=headers, data=payload)
        text = response.text
        data = json.loads(text)
        Store_name = data['name']
        address = data['address']
        print(j, Store_name, address, id)
        j += 1
#datalab의 내용중 context, popularity의 내용을 검색함
        if 'context' in data['datalab']:
            if data['datalab']['context']['goal'] == None:
                pass
            else:
                visite_goal = str(data['datalab']['context']['goal']).strip('[').strip(']')
            if data['datalab']['context']['atmosphere'] == None:
                pass
            else:
                atmosphere = str(data['datalab']['context']['atmosphere']).strip('[').strip(']')
            if data['datalab']['context']['represent'] == None:
                pass
            else:
                represent = str(data['datalab']['context']['represent']).strip('[').strip(']')
            if data['datalab']['context']['topic'] == None:
                pass
            else:
                topic = str(data['datalab']['context']['topic']).strip('[').strip(']')
        else:
            visite_goal = '-'
            atmosphere = '-'
            represent = '-'
            topic = '-'
        if 'popularity' in data['datalab']:
            if data['datalab']['popularity']['age']['values'] == None:
                pass
            elif '10' in data['datalab']['popularity']['age']['values']:
                values = data['datalab']['popularity']['age']['values']
                total = 0
                for k in values.values():
                    total += k
                key1 = round(float(values['10'] / total * 100), 1)
                key2 = round(float(values['20'] / total * 100), 1)
                key3 = round(float(values['30'] / total * 100), 1)
                key4 = round(float(values['40'] / total * 100), 1)
                key5 = round(float(values['50'] / total * 100), 1)
                key6 = round(float(values['60'] / total * 100), 1)
            if data['datalab']['popularity']['gender'] == None:
                pass
            elif data['datalab']['popularity']['gender'] == 'N':
                pass
            else:
                gender = dict(data['datalab']['popularity']['gender'])
                gender1 = round(float(gender['f'] * 100), 1)
                gender2 = round(float(gender['m'] * 100), 1)
        else:
            key1 = 0.0
            key2 = 0.0
            key3 = 0.0
            key4 = 0.0
            key5 = 0.0
            key6 = 0.0
            gender1 = 0.0
            gender2 = 0.0
#동일한 매장이름, 주소, id가 있으면 스킵함
        if [Store_name, address, id] in list_data:
            continue
        else:
            list_data.append([Store_name, address, id, visite_goal, atmosphere, represent, topic, key1, key2, key3, key4, key5, key6, gender1, gender2])
  else:
      continue
columns = ['점포명', '주소', 'ID', '방문목적', '분위기', '시그니처', '인기토픽', '10대', '20대', '30대', '40대', '50대', '60대', '여자', '남자']
real_df = pd.DataFrame(list_data, columns=columns).sort_values('주소', ascending=False)
real_df.head()
real_df.info()
real_df.to_excel('./4 ' + search_area + '지역 ' + search_business + ' datalab-requests' + curr_date + '.xlsx', index=True)
