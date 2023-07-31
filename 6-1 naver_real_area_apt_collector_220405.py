from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from datetime import datetime
import openpyxl
import pandas as pd
import requests
import json
import time
import os

j = 0
sheet_list = []

search_area = input('검색할 곳의 위도와 경도를 입력하세요 : ')
url1 = search_area.split(',')
search_area = input('검색할 지역명을 입력하세요 : ')
leftLon = round(float(url1[1]) - 0.0034332, 7)
rightLon = round(float(url1[1]) + 0.0034332, 7)
topLat = round(float(url1[0]) + 0.0015866, 7)
bottomLat = round(float(url1[0]) - 0.0015866, 7)
curr_date = datetime.strftime(datetime.now(), "%y_%m_%d")

excel_url = '6-1 네이버부동산 APT ' + search_area + ' 매물조회 ' + curr_date + '.xlsx'
try:
    book = openpyxl.load_workbook(excel_url)
except:
    sheet_count = 0
    last_sheet = ''
    pass
else:
    last_sheet = book.worksheets

url = 'https://new.land.naver.com/api/complexes/single-markers/2.0?complexes/single-markers/2.0?' \
      'zoom=17&priceType=RETAIL&realEstateType=APT&tradeType=A1%3AB1%3AB2&oldBuildYears&recentlyBuildYears&minHouseHoldCount&maxHouseHoldCount&directions=' \
      '&leftLon=' + str(leftLon) + '&rightLon=' + str(rightLon) + '&topLat=' + str(topLat) + '&bottomLat=' + str(bottomLat)
payload={}
headers = {
    'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6IlJFQUxFU1RBVEUiLCJpYXQiOjE2NDc0Mzc1MjksImV4cCI6MTY0NzQ0ODMyOX0.BFkee1Qrgl_X5xEQ2iNrFfIu4EZ6VI2scTiW0Q8Rxfo',
    'Accept': "*/*",
    'Accept-Encoding': "gzip, deflate, br",
    'Accept-Language': "ko-KR, ko;q=0.9, en-US;q=0.8, en;q=0.7",
    'Cache-Control': "no-cache",
    'Postman-Token': "adbba748-cb85-4fb4-8f6a-4be441f19cc3",
    'Host': "m.land.naver.com",
    'Connection': "keep-alive",
    'cache-control': "no-cache",
    'User-Agent': "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.104 Whale/3.13.131.36 Safari/537.36"
}
response = requests.request("GET", url, headers=headers, data=payload)
text = response.text
data = json.loads(text)
for list in data:
    list_data = []
    dangi_id = list['markerId']
    name = list['complexName']
    for count in range(len(last_sheet)):
        sheet_list.append(last_sheet[count].title)

    if name in sheet_list:
        continue
    print(dangi_id, name)
    check = input('자료수집을 하시겠습니까?(ex: 예 : 1, 아니오 : 2) : ')
    if check == '1':
        pass
    else:
        continue
    print(dangi_id, name)
    for i in range(10):
        url = 'https://new.land.naver.com/api/articles/complex/' + str(dangi_id) + '?priceType=RETAIL&tradeType=&page=' + str(i+1)
        payload = {}
        headers = {
        'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6IlJFQUxFU1RBVEUiLCJpYXQiOjE2NDczMjA4NDAsImV4cCI6MTY0NzMzMTY0MH0.CVMz-aQokatAj0AXsYi6maU5zvzYTBurOQfhW7XUpnU',
        'User-Agent': "PostmanRuntime/7.20.0",
        'Accept': "*/*",
        'Cache-Control': "no-cache",
        'Postman-Token': "adbba748-cb85-4fb4-8f6a-4be441f19cc3",
        'Host': "m.land.naver.com",
        'Accept-Encoding': "gzip, deflate",
        'Connection': "keep-alive",
        'cache-control': "no-cache"
        }
        response = requests.request("GET", url, headers=headers, data=payload)
        text = response.text
        data1 = json.loads(text)
        if 'articleList' in data1:
            for list in data1['articleList']:
                id = list['articleNo']
                url = 'https://new.land.naver.com/api/articles/' + str(id) + '?complexNo='
                payload = {}
                headers = {
                    'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6IlJFQUxFU1RBVEUiLCJpYXQiOjE2NDczMjA4NDAsImV4cCI6MTY0NzMzMTY0MH0.CVMz-aQokatAj0AXsYi6maU5zvzYTBurOQfhW7XUpnU',
                    'User-Agent': "PostmanRuntime/7.20.0",
                    'Accept': "*/*",
                    'Cache-Control': "no-cache",
                    'Postman-Token': "adbba748-cb85-4fb4-8f6a-4be441f19cc3",
                    'Host': "m.land.naver.com",
                    'Accept-Encoding': "gzip, deflate",
                    'Connection': "keep-alive",
                    'cache-control': "no-cache"
                }
                response = requests.request("GET", url, headers=headers, data=payload)
                text = response.text
                data2 = json.loads(text)
                articleNo = data2['articleAddition']['articleNo']
                try:
                    articleName = data2['articleAddition']['articleName']
                except:
                    articleName = "-"
                    pass
                try:
                    articleConfirmYmd = data2['articleAddition']['articleConfirmYmd']
                except:
                    articleConfirmYmd = "-"
                    pass
                try:
                    floorInfo = data2['articleAddition']['floorInfo']
                except:
                    floorInfo = '-'
                    pass
                try:
                    exclusiveSpace = data2['articleSpace']['exclusiveSpace']
                except:
                    exclusiveSpace = 0.0
                    pass
                try:
                    supplySpace = data2['articleSpace']['supplySpace']
                except:
                    supplySpace = 0.0
                    pass
                try:
                    cpPcArticleUrl = data2['articleAddition']['cpPcArticleUrl']
                except:
                    cpPcArticleUrl = "-"
                    pass
                try:
                    direction = data2['articleAddition']['direction']
                except:
                    direction = "-"
                    pass
                try:
                    buildingUseAprvYmd = data2['articleFacility']['buildingUseAprvYmd']
                except:
                    buildingUseAprvYmd = "-"
                    pass
                try:
                    dealPrice = data2['articlePrice']['dealPrice']
                except:
                    dealPrice = 0
                    pass
                try:
                    warrantPrice = data2['articlePrice']['warrantPrice']
                except:
                    warrantPrice = 0
                    pass
                try:
                    rentPrice = data2['articlePrice']['rentPrice']
                except:
                    rentPrice = 0
                    pass
                try:
                    realtorName = data2['articleRealtor']['realtorName']
                except:
                    realtorName = "-"
                    pass
                try:
                    representativeName = data2['articleRealtor']['representativeName']
                except:
                    representativeName = "-"
                    pass
                try:
                    cellPhoneNo = data2['articleRealtor']['cellPhoneNo']
                except:
                    cellPhoneNo = "-"
                    pass
                try:
                    articleFeatureDescription = data2['articleDetail']['articleFeatureDescription'].strip('(').strip(')')
                except:
                    articleFeatureDescription = "-"
                    pass
                try:
                    dongNm = data2['landPrice']['dongNm']
                except:
                    dongNm = '-'
                    pass
                try:
                    hoNm = data2['landPrice']['hoNm']
                except:
                    hoNm = '-'
                    pass
                print(j)
                j += 1
                if j % 70 == 0:
                    time.sleep(6)
                list_data.append([articleNo, articleName, dongNm, hoNm, exclusiveSpace, supplySpace, floorInfo, dealPrice, warrantPrice, rentPrice, direction, articleFeatureDescription, \
                                  articleConfirmYmd, realtorName, representativeName, cellPhoneNo, cpPcArticleUrl])

    columns = ['매물번호', '건물명', '동', '호', '전용면적', '계약면적', '층수', '매매가', '보증금', '월세', '방향', '특징', '확인일자', '중개사무소', '중개인', '전화', '블로그']
    real_df = pd.DataFrame(list_data, columns=columns).sort_values(by='건물명', ascending=False)
    if not os.path.exists('./6-1 네이버부동산 APT ' + search_area + ' 매물조회 ' + curr_date + '.xlsx'):
        with pd.ExcelWriter('./6-1 네이버부동산 APT ' + search_area + ' 매물조회 ' + curr_date + '.xlsx', mode='w', engine='openpyxl') as writer:
            real_df.to_excel(writer, sheet_name=f"{name}")
    else:
        with pd.ExcelWriter('./6-1 네이버부동산 APT ' + search_area + ' 매물조회 ' + curr_date + '.xlsx', mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
            real_df.to_excel(writer, sheet_name=f"{name}")
    # time.sleep(10)

#모든 쉬트를 읽어와서 하나로 합치기
columns = ['매물번호', '건물명', '동', '호', '전용면적', '계약면적', '층수', '매매가', '보증금', '월세', '방향', '특징', '확인일자', '중개사무소', '중개인', '전화', '블로그']
df_all = pd.read_excel(excel_url, sheet_name = None)
concatted_df = pd.concat(df_all, ignore_index=False)
real_df2 = pd.DataFrame(concatted_df, columns=columns).sort_values(by='건물명', ascending=False)
with pd.ExcelWriter('./6-1 네이버부동산 APT ' + search_area + ' 매물조회 ' + curr_date + '.xlsx', mode='a', engine='openpyxl') as writer:
    real_df2.to_excel(writer, sheet_name=f"{curr_date}")