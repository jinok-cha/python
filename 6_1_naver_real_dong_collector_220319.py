from datetime import datetime
import pandas as pd
import requests
import json
import openpyxl
import time

j = 0
list_data = []

search_area = input('검색할 동을 입력하세요 : ')
search_area = search_area.split(' ')
sido = search_area[0]
gungu = search_area[1]
dong = search_area[2]
type = input('검색할 물건을 선택하세요(SG:SMS:GTCG:APTHGJ:GM:TJ, 엔터는 건물임) : ')
if type == '':
    type = 'GM'
# macket_rate = input('시장환원율을 입력하세요 : ')
curr_date = datetime.strftime(datetime.now(), "%y_%m_%d")

def get_sido_info():
    temp1 = []
    down_url = 'https://new.land.naver.com/api/regions/list?cortarNo=0000000000'
    r = requests.get(down_url,data={"sameAddressGroup":"false"},headers={
        "Accept-Encoding": "gzip",
        "Host": "new.land.naver.com",
        "Referer": "https://new.land.naver.com/complexes/102378?ms=37.5018495,127.0438028,16&a=APT&b=A1&e=RETAIL",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36"
    })
    r.encoding = "utf-8-sig"
    temp=json.loads(r.text)
    for i in range(len(temp['regionList'])):
        temp1.append([temp['regionList'][i]["cortarNo"], temp['regionList'][i]["cortarName"]])
    return temp1

def get_gungu_info(sido_code):
    temp2 = []
    down_url = 'https://new.land.naver.com/api/regions/list?cortarNo='+sido_code
    r = requests.get(down_url,data={"sameAddressGroup":"false"},headers={
        "Accept-Encoding": "gzip",
        "Host": "new.land.naver.com",
        "Referer": "https://new.land.naver.com/complexes/102378?ms=37.5018495,127.0438028,16&a=APT&b=A1&e=RETAIL",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36"
    })
    r.encoding = "utf-8-sig"
    temp=json.loads(r.text)
    for i in range(len(temp['regionList'])):
        temp2.append([temp['regionList'][i]["cortarNo"], temp['regionList'][i]["cortarName"]])
    return temp2

def get_dong_info(gungu_code):
    temp3 = []
    for gungu in gungu_code:
        down_url = 'https://new.land.naver.com/api/regions/list?cortarNo='+gungu_code
        r = requests.get(down_url,data={"sameAddressGroup":"false"},headers={
            "Accept-Encoding": "gzip",
            "Host": "new.land.naver.com",
            "Referer": "https://new.land.naver.com/complexes/102378?ms=37.5018495,127.0438028,16&a=APT&b=A1&e=RETAIL",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36"
        })
        r.encoding = "utf-8-sig"
        temp = json.loads(r.text)
        for i in range(len(temp['regionList'])):
            temp3.append([temp['regionList'][i]["cortarNo"], temp['regionList'][i]["cortarName"]])
    return temp3

sido_list = get_sido_info()
for sidos in sido_list:
    if sido == sidos[1]:
        sido_code = sidos[0]
        break
gungu_list = get_gungu_info(sido_code)
for gungus in gungu_list:
    if gungu == gungus[1]:
        gungu_code = gungus[0]
        break
dong_list = get_dong_info(gungu_code)
for dongs in dong_list:
    if dong == dongs[1]:
        dong_code = dongs[0]
        break
for i in range(10):
    url = 'https://new.land.naver.com/api/articles?cortarNo=' + dong_code + '&order=rank&realEstateType=' + type + '&priceType=RETAIL&page=' + str(i+1)
    payload={}
    headers = {
        'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6IlJFQUxFU1RBVEUiLCJpYXQiOjE2NDcxNzg2NTQsImV4cCI6MTY0NzE4OTQ1NH0.jxNhCqLWDDvHPvF8KbqULExCnEujbv67wGpws4ibK-U',
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
    data = json.loads(text)
    if 'articleList' in data:
        for list in data['articleList']:
            id = list['articleNo']
            url = 'https://new.land.naver.com/api/articles/' + str(id) + '?complexNo='
            payload = {}
            headers = {    'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6IlJFQUxFU1RBVEUiLCJpYXQiOjE2NDcxNzg2NTQsImV4cCI6MTY0NzE4OTQ1NH0.jxNhCqLWDDvHPvF8KbqULExCnEujbv67wGpws4ibK-U',
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

            articleNo = data1['articleDetail']['articleNo']
            try:
                noneAptBuildingName = data1['articleDetail']['noneAptBuildingName']
            except:
                noneAptBuildingName = " "
                pass
            try:
                articleConfirmYMD = data1['articleDetail']['articleConfirmYMD']
            except:
                articleConfirmYMD = " "
                pass
            try:
                lawUsage = data1['articleDetail']['lawUsage']
            except:
                lawUsage = " "
                pass
            try:
                floorInfo = data1['articleAddition']['floorInfo']
            except:
                floorInfo = " "
                pass
            try:
                area1 = data1['articleAddition']['area1']
            except:
                area1 = 0.0
                pass
            try:
                area2 = data1['articleAddition']['area2']
            except:
                area2 = 0.0
                pass
            try:
                cpPcArticleUrl = data1['articleAddition']['cpPcArticleUrl']
            except:
                cpPcArticleUrl = " "
                pass
            try:
                directionTypeName = data1['articleFacility']['directionTypeName']
            except:
                directionTypeName = " "
                pass
            try:
                usageAreaTypeName = data1['articleFacility']['usageAreaTypeName']
            except:
                usageAreaTypeName = " "
                pass
            try:
                buildingUseAprvYmd = data1['articleFacility']['buildingUseAprvYmd']
            except:
                buildingUseAprvYmd = " "
                pass
            try:
                dealPrice = data1['articlePrice']['dealPrice']
            except:
                dealPrice = 0
                pass
            try:
                allWarrantPrice = data1['articlePrice']['allWarrantPrice']
            except:
                allWarrantPrice = 0
                pass
            try:
                allRentPrice = data1['articlePrice']['allRentPrice']
            except:
                allRentPrice = 0
                pass
            try:
                rate = (allRentPrice * 12) / (dealPrice - allWarrantPrice) * 100
            except:
                rate = 0
                pass
            try:
                priceBySpace = data1['articlePrice']['priceBySpace']
            except:
                priceBySpace = 0
                pass
            try:
                realtorName = data1['articleRealtor']['realtorName']
            except:
                realtorName = " "
                pass
            try:
                representativeName = data1['articleRealtor']['representativeName']
            except:
                representativeName = " "
                pass
            try:
                cellPhoneNo = data1['articleRealtor']['cellPhoneNo']
            except:
                cellPhoneNo = " "
                pass
            try:
                etcParkInfo = data1['articleBuildingRegister']['etcParkInfo'].strip('(').strip(')')
            except:
                etcParkInfo = " "
                pass
            try:
                bcRat = data1['articleBuildingRegister']['bcRat']
            except:
                bcRat = 0.0
                pass
            try:
                vlRat = data1['articleBuildingRegister']['vlRat']
            except:
                vlRat = 0.0
                pass
            try:
                newPlatPlc = data1['articleBuildingRegister']['newPlatPlc']
            except:
                newPlatPlc = 0.0
                pass
            print(j)
            j += 1
            list_data.append([articleNo, dealPrice, priceBySpace, noneAptBuildingName, area1, area2, bcRat, vlRat, usageAreaTypeName, lawUsage, floorInfo, directionTypeName, etcParkInfo, \
                              buildingUseAprvYmd, allWarrantPrice, allRentPrice, round(float(rate), 2), newPlatPlc, articleConfirmYMD, realtorName, representativeName, cellPhoneNo, cpPcArticleUrl])
columns = ['매물번호', '매매가격', '평단가', '건물명', '대지면적', '연면적', '건폐율', '용적률', '토지용도', '건물용도', '층수', '방향', '주차', '사용승인일자', '보증금', '월세', '환원율%', '물건주소', '확인일자', '중개사무소', '중개인', '전화', '블로그']
real_df = pd.DataFrame(list_data, columns=columns).sort_values(by='대지면적', ascending=False)
real_df.head()
real_df.info()
real_df.to_excel('./6-1 ' + dong + '지역 네이버부동산 ' + type + ' 매물조회 ' + curr_date + '.xlsx', sheet_name=f"{dong}", index=False)

