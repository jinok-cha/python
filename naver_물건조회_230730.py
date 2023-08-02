import requests
import time
import pandas as pd
from tkinter import *
import tkinter.ttk
import tkinter.messagebox as msgbox
import math
from bs4 import BeautifulSoup
import json
from openpyxl import Workbook
import datetime


ws = []
wb = Workbook()
def btnsearchcmd():
    maximum_count = 0
    keyword = entry.get()

    url = "https://m.land.naver.com/map/37.4979304:127.0273088:18:/SG:SMS:GM/A1:B1:B2"#"https://m.land.naver.com/search/result/{}".format(keyword)
    res = requests.get(url, headers={'user-agent': 'Mozilla/5.0'})
    res.raise_for_status()

    today = datetime.datetime.now()
    curr_date = today.strftime('%Y-%m-%d')

    soup = (str)(BeautifulSoup(res.text, "lxml"))
    value = soup.split("filter: {")[1].split("}")[0].replace(" ", "").replace("'", "")
    print(value)
    lat = value.split("lat:")[1].split(",")[0]
    lon = value.split("lon:")[1].split(",")[0]
    z = value.split("z:")[1].split(",")[0]
    cortarNo = value.split("cortarNo:")[1].split(",")[0]
    rletTpCds = value.split("rletTpCds:")[1].split(",")[0]
    tradTpCds = value.split("tradTpCds:")[1].split()[0]

    # lat - btm : 37.550985 - 37.4331698 = 0.1178152
    # top - lat : 37.6686142 - 37.550985 = 0.1176292
    lat_margin = 0.118

    # lon - lft : 126.849534 - 126.7389841 = 0.1105499
    # rgt - lon : 126.9600839 - 126.849534 = 0.1105499
    lon_margin = 0.111

    btm = float(lat) - lat_margin
    lft = float(lon) - lon_margin
    top = float(lat) + lat_margin
    rgt = float(lon) + lon_margin

    # 최초 요청 시 디폴트 값으로 설정되어 있으나, 원하는 값으로 구성
    z = 15
    rletTpCds = "GM"  # 상가
    tradTpCds = "A1:B1:B2"  # 매매/전세/월세 매물 확인

    # clusterList?view 를 통한 그룹(단지)의 데이터를 가져온다.
    remaked_URL = "https://m.land.naver.com/cluster/clusterList?view=atcl&cortarNo={}&rletTpCd={}&tradTpCd={}&z={}&lat={}&lon={}&btm={}&lft={}&top={}&rgt={}&addon=COMPLEX&bAddon=COMPLEX&isOnlyIsale=false" \
        .format(cortarNo, rletTpCds, tradTpCds, z, lat, lon, btm, lft, top, rgt)
    res2 = requests.get(remaked_URL, headers={'user-agent': 'Mozilla/5.0'})
    json_str = json.loads(json.dumps(res2.json()))
    values = json_str['data']['ARTICLE']

    # 큰 원으로 구성되어 있는 전체 매물그룹(values)을 load 하여 한 그룹씩 세부 쿼리 진행
    for v in values[:10]:
        lgeo = v['lgeo']
        count = v['count']
        z2 = v['z']
        lat2 = v['lat']
        lon2 = v['lon']
        len_pages = count / 20 + 1

        for idx in range(1, math.ceil(len_pages)):
            remaked_URL2 = "https://m.land.naver.com/cluster/ajax/articleList?""itemId={}&mapKey=&lgeo={}&showR0=&" \
                           "rletTpCd={}&tradTpCd={}&z={}&lat={}&""lon={}&totCnt={}&cortarNo={}&page={}" \
                .format(lgeo, lgeo, rletTpCds, tradTpCds, z2, lat2, lon2, count, cortarNo, idx)
            res3 = requests.get(remaked_URL2, headers={'user-agent': 'Mozilla/5.0'})
            json_str1 = json.loads(json.dumps(res3.json()))
            atclNo = json_str1['body'][0]['atclNo']  # 물건번호
            rletTpNm = json_str1['body'][0]['rletTpNm']  # 상가구분
            tradTpNm = json_str1['body'][0]['tradTpNm']  # 매매/전세/월세 구분
            prc = json_str1['body'][0]['prc']  # 가격
            spc1 = json_str1['body'][0]['spc1']  # 계약면적(m2) -> 평으로 계산 : * 0.3025
            spc2 = json_str1['body'][0]['spc2']  # 전용면적(m2) -> 평으로 계산 : * 0.3025
            hanPrc = json_str1['body'][0]['hanPrc']  # 보증금
            rentPrc = json_str1['body'][0]['rentPrc']  # 월세
            flrInfo = json_str1['body'][0]['flrInfo']  # 층수(물건층/전체층)
            lat1 = json_str1['body'][0]['lat']  # 위도
            lng2 = json_str1['body'][0]['lng']  # 경도
            tagList = json_str1['body'][0]['tagList']  # 기타 정보
            rltrNm = json_str1['body'][0]['rltrNm']  # 부동산
            detaild_information = "https://m.land.naver.com/article/info/{}".format(atclNo)

            # 표에 삽입될 데이터
            tablelist = [str(rletTpNm), str(tradTpNm), str(format(prc, ',')), str(spc1),
                         str(spc2), str(hanPrc), str(format(rentPrc, ',')), str(flrInfo),
                         str(lat1), str(lng2), str(tagList), str(rltrNm), detaild_information]

            tableview.insert("", 'end', values=tablelist)

    # 엑셀시트에 데이터 append
    ws.append([str(rletTpNm), str(tradTpNm), str(prc), str(spc1),
               str(spc2), str(hanPrc), str(rentPrc), str(flrInfo),
               str(lat1), str(lng2), str(tagList), str(rltrNm), detaild_information])

    # 검색 완료 후 엑셀 저장 버튼 활성화
    btn_exportexcel.config(state="active")

    # 엑셀 저장 후 프로그램 종료 버튼 활성화
    btn_exit.config(state="active")

def btnexportexcel():
    now = datetime.datetime.now()
    nowDatetime = now.strftime('%Y%m%d_%H%M%S')
    keyword = entry.get()

    file_name = keyword + "_" + nowDatetime + ".xlsx"
    wb.save(filename="./" + file_name)
    msgbox.showinfo("파일 저장", "'" + file_name + "' 파일로 정상적으로 추출되었습니다.")

def btnexit():
    exit()

root = Tk()
root.title("부동산 상가 매물 검색 프로그램")

#검색 프레임 (entry, 검색버튼, 엑셀 버튼)
search_frame = Frame(root)
search_frame.pack(expand=True, pady=10,fill="both")

#검색 입력 창
entry = Entry(search_frame)
entry.pack(side="left",fill="both", expand=True)
entry.focus()
entry.configure(state='normal')

#entry에 클릭했을 때 on_forcus_in 함수 실행
x_focus_in = entry.bind('<Button-1>', lambda x: entry) #<Button-1> 왼쪽버튼 클릭
x_focus_out = entry.bind('<FocusOut>', lambda x: entry) #<FocusOut> 위젯선택 풀릴 시 (다른 곳 클릭 or tab)

#검색버튼
btn_search = Button(search_frame, text="검색", padx=5, pady=5, command = btnsearchcmd)
btn_search.pack(side="left", padx=5, fill="both")

#엑셀 저장 버튼
btn_exportexcel = Button(search_frame, text="엑셀 저장",  padx=5, pady=5, command = btnexportexcel, state=DISABLED)
btn_exportexcel.pack(side="left", padx=5,fill="both")

#프로그램 종료 버튼
btn_exit = Button(search_frame,  text="프로그램 종료", padx=5, pady=5, command=btnexit)
btn_exit.pack(side="left", padx=5,fill="both")

# 상가 구분 프레임
sg_condition_frame = Frame(root)
sg_condition_frame.pack(side="top", pady=20,fill="both")

# 큰 프레인 안에 좌:"상가 구분", 우:"거래 유형" 으로 레이아웃 쪼개기
# LableFrame 을 활용하여, checkbutton 을 묶어서 제목:"상가 구분" 붙이기
frame_middle_left = LabelFrame(sg_condition_frame, text="상가 구분")
frame_middle_left.pack(side="left", fill="both", expand=True)

sg_chk1 = IntVar()
sg_chk1_box = Checkbutton(frame_middle_left, text="상가", variable = sg_chk1)# 상가
sg_chk1_box.pack(side="left")

sg_chk2 = IntVar()
sg_chk2_box = Checkbutton(frame_middle_left, text="상가주택", variable = sg_chk2)# 상가주택
sg_chk2_box.pack(side="left")

sg_chk3 = IntVar()
sg_chk3_box = Checkbutton(frame_middle_left, text="건물", variable = sg_chk3)# 건물
sg_chk3_box.pack(side="left")
sg_chk3_box.select()

# 거래 유형 프레임
frame_middle_right = LabelFrame(sg_condition_frame, text="거래유형")
frame_middle_right.pack(side="right", fill="both", expand=True)

tr_type1 = IntVar()
tr_type1_box = Checkbutton(frame_middle_right, text="매매", variable = tr_type1)# 매매
tr_type1_box.pack(side="left")
tr_type1_box.select()

tr_type2 = IntVar()
tr_type2_box = Checkbutton(frame_middle_right, text="월세", variable = tr_type2)# 월세
tr_type2_box.pack(side="left")

# 결과 출력 프레임
result_print_frame = LabelFrame(root, text = "검색 결과")
result_print_frame.pack(side="top", fill="both")

list_frame  = Frame(result_print_frame)
list_frame.pack(side="top", fill="both")

scrollbar = Scrollbar(list_frame)
scrollbar.pack(side="right", fill = "y")

tableview = tkinter.ttk.Treeview(list_frame, columns=["rletTpNm","tradTpNm","prc","spc1","spc2","hanPrc",
                     "rentPrc", "flrInfo", "lat1", "lng2", "tagList", "rltrNm", "detaild_information"],
                     displaycolumns=["rletTpNm","tradTpNm","prc","spc1","spc2","hanPrc",
                     "rentPrc", "flrInfo", "lat1", "lng2", "tagList", "rltrNm", "detaild_information"],
                     height=40, yscrollcommand=scrollbar.set)
tableview.pack(fill="both")

# 각 컬럼 설정. 컬럼 이름, 컬럼 넓이, 정렬 등
tableview.column("rletTpNm", width=80, anchor="center")
tableview.heading("rletTpNm", text="물건 구분")

tableview.column("tradTpNm", width=80, anchor="center")
tableview.heading("tradTpNm", text="거래 유형", anchor="center")

tableview.column("prc", width=80, anchor="center")
tableview.heading("prc", text="가격", anchor="center")

tableview.column("spc1", width=80, anchor="center")
tableview.heading("spc1", text="계약면적", anchor="center")

tableview.column("spc2", width=80, anchor="center")
tableview.heading("spc2", text="전용면적", anchor="center")

tableview.column("hanPrc", width=80, anchor="center")
tableview.heading("hanPrc", text="보증금", anchor="center")

tableview.column("rentPrc", width=80, anchor="center")
tableview.heading("rentPrc", text="월세", anchor="center")

tableview.column("flrInfo", width=80, anchor="center")
tableview.heading("flrInfo", text="층수", anchor="center")

tableview.column("lat1", width=100, anchor="center")
tableview.heading("lat1", text="위도", anchor="center")

tableview.column("lng2", width=100, anchor="center")
tableview.heading("lng2", text="경도", anchor="center")

tableview.column("tagList", width=200, anchor="center")
tableview.heading("tagList", text="기타사항", anchor="center")

tableview.column("rltrNm", width=200, anchor="center")
tableview.heading("rltrNm", text="중개사", anchor="center")

tableview.column("detaild_information", width=300, anchor="center")
tableview.heading("detaild_information", text="비고", anchor="center")

#스크롤바를 움직일 때 표도 같이 이동할 수 있도록 적용
scrollbar.config(command=tableview.yview)

root.mainloop()