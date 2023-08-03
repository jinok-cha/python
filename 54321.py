import requests
import time
import pandas as pd
import math
from bs4 import BeautifulSoup
import json
import datetime
import os

list_data = []

keyword = "12345"#input("조회할 지역을 입력하세요 : ")
today = datetime.datetime.now()
curr_date = today.strftime('%Y-%m-%d')

url = "https://m.land.naver.com/map/37.497503:127.027438:18/SG:SMS:GM/A1:B1:B2" #input("검색할 지역의 url을 입력하세요 : ")
res = requests.get(url, headers={'user-agent': 'Mozilla/5.0'})
res.raise_for_status()
soup = str(BeautifulSoup(res.text, "lxml"))
value = soup.split("filter: {")[1].split("}")[0].replace(" ", "").replace("'", "")
lat = value.split("lat:")[1].split(",")[0]
lon = value.split("lon:")[1].split(",")[0]
z = value.split("z:")[1].split(",")[0]
#cortarNo = "2635010600" #input("cortarNo를 입력하세요  : ")
rletTpCds = value.split("rletTpCds:")[1].split(",")[0]
tradTpCds = value.split("tradTpCds:")[1].split()[0]

# lat - btm : 37.550985 - 37.4331698 = 0.1178152
# top - lat : 37.6686142 - 37.550985 = 0.1176292
lat_margin = 0.0025#0.018
# lon - lft : 126.849534 - 126.7389841 = 0.1105499
# rgt - lon : 126.9600839 - 126.849534 = 0.1105499
lon_margin = 0.0025 #0.111
btm = float(lat) - lat_margin
lft = float(lon) - lon_margin
top = float(lat) + lat_margin
rgt = float(lon) + lon_margin

# clusterList?view 를 통한 그룹(단지)의 데이터를 가져온다.
remaked_URL =f"https://m.land.naver.com/cluster/clusterList?view=atcl&rletTpCd={rletTpCds}&tradTpCd={tradTpCds}&z={z}&lat={lat}&lon={lon}&btm={btm}&lft={lft}&top={top}&rgt={rgt}&addon=COMPLEX&bAddon=COMPLEX&isOnlyIsale=false"
res2 = requests.get(remaked_URL, headers={'user-agent': 'Mozilla/5.0'})
json_str = json.loads(json.dumps(res2.json()))
values = json_str['data']['ARTICLE']

# 큰 원으로 구성되어 있는 전체 매물그룹(values)을 load 하여 한 그룹씩 세부 쿼리 진행
for v in values[:1]:
    lgeo = v['lgeo']
    count = v['count']
    z2 = v['z']
    lat2 = v['lat']
    lon2 = v['lon']
    len_pages = count / 20 + 1

    for idx in range(1, math.ceil(len_pages)):
        remaked_URL2 = f"https://m.land.naver.com/cluster/ajax/articleList?itemId={lgeo}&mapKey=&lgeo={lgeo}&showR0=&rletTpCd={rletTpCds}&tradTpCd={tradTpCds}&z={z2}&lat={lat2}&lon={lon}&totCnt={count}&page={idx}"
        res3 = requests.get(remaked_URL2, headers={'user-agent': 'Mozilla/5.0'})
        json_str1 = json.loads(json.dumps(res3.json()))
        for i in range(len(json_str1['body'])):
            print(idx*20+i, count)
            try:
                atclNo = json_str1['body'][i]['atclNo']  # 물건번호
                rletTpNm = json_str1['body'][i]['rletTpNm']  # 상가구분
                tradTpNm = json_str1['body'][i]['tradTpNm']  # 매매/전세/월세 구분
                prc = json_str1['body'][i]['prc']  # 가격
                spc1 = round(float(json_str1['body'][i]['spc1']) * 0.3025, 2)   # 계약면적(m2) -> 평으로 계산 : * 0.3025
                spc2 = round(float(json_str1['body'][i]['spc2']) * 0.3025, 2)   # 전용면적(m2) -> 평으로 계산 : * 0.3025
                hanPrc = json_str1['body'][i]['hanPrc'].replace(",","").replace("억 ", "").replace("억","0000")  # 보증금
                rentPrc = json_str1['body'][i]['rentPrc']  # 월세
                flrInfo = json_str1['body'][i]['flrInfo'].split("/")  # 층수(물건층/전체층)
                flrInfo[0] = int(flrInfo[0].replace("B", "-"))
                flrInfo[1] = flrInfo[1].replace("B", "-")
                lat1 = json_str1['body'][i]['lat']  # 위도
                lng2 = json_str1['body'][i]['lng']  # 경도
                tagList = str(json_str1['body'][i]['tagList']).replace("[", "").replace("]", "").replace("'", "")  # 기타 정보
                rltrNm = json_str1['body'][i]['rltrNm']  # 부동산
                detaild_information = "https://m.land.naver.com/article/info/{}".format(atclNo)
            except:
                atclNo = 0
                rletTpNm = 0
                tradTpNm = 0
                prc = 0
                spc1 = 0
                spc2 = 0
                hanPrc = 0
                rentPrc = 0
                flrInfo = [0, 0]
                lat1 = 0
                lng2 = 0
                tagList = 0
                rltrNm = 0
                detaild_information = 0
            try:
                avg_hanPrc = round(float(hanPrc.replace(",", "").replace("억", "0000").replace("억 ", "")) / spc2, 1)
                avg_rentPrc = round(rentPrc / spc2, 1)
            except:
                avg_hanPrc = 0
                avg_rentPrc = 0
            list_data.append([atclNo, rletTpNm, tradTpNm, prc, spc1, spc2, int(hanPrc), rentPrc, avg_hanPrc, avg_rentPrc, flrInfo[0], flrInfo[1],
               lat1, lng2, tagList, rltrNm, detaild_information])
            time.sleep(2)
        time.sleep(10)
columns = ['물건번호', '상가구분', '거래방식', '매매가', '계약면적(평)', '전용면적(평)', '보증금', '월세', '평당보증금', '평당월세', '해당층수', '총층수', '위도', '경도', '기타정보', '부동산', '비고']
real_df = pd.DataFrame(list_data, columns=columns).sort_values(by='평당월세', ascending=False)
if not os.path.exists('./' + keyword + '_' + curr_date + '.xlsx'):
    with pd.ExcelWriter('./' + keyword + '_' +curr_date + '.xlsx', mode='w', engine='openpyxl') as writer:
        real_df.to_excel(writer, sheet_name=f"{keyword}")
else:
    with pd.ExcelWriter('./' + keyword + '_' + curr_date + '.xlsx', mode='a', engine='openpyxl') as writer:
        real_df.to_excel(writer, sheet_name=f"{keyword}")