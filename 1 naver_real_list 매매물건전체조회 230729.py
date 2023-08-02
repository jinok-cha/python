from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
import pandas as pd
from datetime import datetime
import openpyxl

filename = './pusan_subway_st.xlsx'
book = openpyxl.load_workbook(filename)
worksheet_num = input('전체 1, 중심상권 2, 테스트 3 : ')
sheet = book.worksheets[int(worksheet_num)]
data = []
real_list = []
core_list = []
for row in sheet.rows:
    data.append([
        row[1].value,
        row[2].value,
        row[3].value
                 ])
curr_date = datetime.strftime(datetime.now(), "%y_%m_%d")
curr_year = 2000 + int(curr_date.split('_')[0])

min_price = int(input('최저가격을 입력하세요(억) : ')) * 10000
max_price = int(input('최고가격을 입력하세요(억) : ')) * 10000

option = webdriver.ChromeOptions()
option.add_argument("start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=option)

for num in range(len(data)-1):
    st_name = data[num+1][0]
    lati = data[num+1][1]
    logi = data[num+1][2]

    url = 'https://new.land.naver.com/offices?ms=' + str(lati) +',' + str(logi) + ',16&a=GM&b=A1&e=RETAIL&ad=true'
    driver.get(url)
    time.sleep(2)

    body = driver.find_element(By.CSS_SELECTOR, '#listContents1 > div > div > div:nth-child(1) > div:nth-child(1) > div > a')
    for i in range(30):
        body.send_keys(Keys.PAGE_DOWN)
    time.sleep(1)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    lists = soup.select("div.item")

    i = 0
    for list in lists:
        price = list.select_one("div > div > a > div.price_line > span.price")
        try:
            if price.text.find('억') > 0:
                sale_price_list = price.text.replace(',','').split('억')
                if sale_price_list[1] == '':
                    sale_price = int(sale_price_list[0]) * 10000
                else:
                    sale_price = int(sale_price_list[0]) * 10000 + int(sale_price_list[1])
            else:
                continue
        except:
            price = 0
        if sale_price < min_price or sale_price > max_price:
            i += 1
            continue
        i += 1
        try:
            body = driver.find_element(By.CSS_SELECTOR,'#listContents1 > div > div > div > div:nth-child(' + str(i) + ') > div > a')
            body.click()
            time.sleep(1)
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            memo = soup.select_one('#detailContents1 > div.detail_box--summary > table > tbody > tr:nth-child(2) > td')
            if memo == None:
                memo = '---'
            else:
                memo = memo.text
            area1 = soup.select_one('#detailContents1 > div.detail_box--summary > table > tbody > tr:nth-child(3) > td')
            if area1 == None:
                area1 = [0,0]
            else:
                area1 = area1.text.replace('㎡','').replace('-', '0.0').split('/') # area1[0] 대지면적, area1[1] 연면적
                area1[0] = float(area1[0])
                area1[1] = float(area1[1])
            area2 = soup.select_one('#detailContents1 > div.detail_box--summary > table > tbody > tr:nth-child(4) > td')
            if area2 == None:
                area2 = [0,0]
            else:
                area2 = area2.text.replace('㎡','').replace('-', '0.0').split('/') # area2[0] 건축면적, area2[1] 전용면적
                area2[0] = float(area2[0])
                area2[1] = float(area2[1])
            try:
                chung = soup.select_one('#detailContents1 > div.detail_box--summary > table > tbody > tr:nth-child(5) > td:nth-child(2)').text
            except:
                chung = 0
            try:
                debt = soup.select_one('#detailContents1 > div.detail_box--summary > table > tbody > tr:nth-child(7) > td:nth-child(2)').text.replace('-','0').replace('만원','').replace(',','').replace('원', '')
                if debt.find('억') > 0:
                    debt_list = debt.split('억')
                    if debt_list[1] == '':
                        debt = int(debt_list[0]) * 10000
                    else:
                        debt = int(debt_list[0]) * 10000 + int(debt_list[1])
                else:
                    if "시세 대비" or "없음" in debt:
                        debt = 0
                    else:
                        debt = int(debt)
            except:
                debt = 0
            try:
                bo_mon = soup.select_one('#detailContents1 > div.detail_box--summary > table > tbody > tr:nth-child(7) > td:nth-child(4)').text.replace('-','0').replace('만원','').replace(',','').replace('원', '').split('/')
                if bo_mon[0].find('억') > 0:
                    bo_mon_list1 = bo_mon[0].split('억')
                    if bo_mon_list1[1] == '':
                        bo_mon[0] = int(bo_mon_list1[0]) * 10000
                    else:
                        bo_mon[0] = int(bo_mon_list1[0]) * 10000 + int(bo_mon_list1[1])
                else:
                    bo_mon[0] = int(bo_mon[0])
                if bo_mon[1].find('억') > 0:
                    bo_mon_list2 = bo_mon[1].split('억')
                    if bo_mon_list2[1] == '':
                        bo_mon[1] = int(bo_mon_list2[0]) * 10000
                    else:
                        bo_mon[1] = int(bo_mon_list2[0]) * 10000 + int(bo_mon_list2[1]) # bo_mon[0] 보증금, bo_mon[1] 월세
                else:
                    bo_mon[1] = int(bo_mon[1])
                if bo_mon[1] != 0:
                    property_value1 = round(float((bo_mon[1] / 0.045) * 12 + bo_mon[0]), -3)
                    property_value2 = round(float((bo_mon[1] / 0.04) * 12 + bo_mon[0]), -3)
                else:
                    property_value1 = 0
                    property_value2 = 0
                if property_value1 == 0:
                    check_list = '??'
                elif sale_price - property_value1 < 0:
                    check_list = '우선확인'
                elif sale_price - property_value2 < 0:
                    check_list = '차선확인'
                else:
                    check_list = '-'
            except:
                    bo_mon = [0,0]
                    property_value1 = 0
                    property_value2 = 0
                    check_list = '-'
            try:
                jungong_date = soup.select_one('#detailContents1 > div.detail_box--summary > table > tbody > tr:nth-child(13) > td:nth-child(2)').text
                if jungong_date == '-':
                    jungong_date = '0000.00.00'
                    struc_price = 0
                    land_price = 0
                    p_land_price = 0
                else:
                    jungong_year = jungong_date.split('.')[0]
                    struc_price = round(int(area1[1] / 3.3 * 400 * (1 - (curr_year - int(jungong_year)) / 40)), -3)
                    land_price = round(int(sale_price - struc_price), -3)
                    p_land_price = round(int(land_price / (area1[0] / 3.3)), -3)
            except:
                jungong_date = 0
                struc_price = 0
                land_price = 0
                p_land_price = 0
            try:
                sale_number = soup.select_one('#detailContents1 > div.detail_box--summary > table > tbody > tr:nth-child(12) > td:nth-child(4)').text
            except:
                sale_number = 0
            try:
                agent_name = soup.select_one('#detailContents1 > div.detail_box--summary > table > tbody > tr > td > div > div.info_agent_title > strong').text
            except:
                agent_name = "-"
            try:
                agent_phone = soup.select_one('#detailContents1 > div.detail_box--summary > table > tbody > tr > td > div > div.info_agent_wrap > dl:nth-child(2) > dd', {'class':'text text--number'}).text
            except:
                agent_phone = 0
            if i > len(lists):
                break
            real_list.append([st_name, sale_price, debt, bo_mon[0], bo_mon[1], property_value1, check_list, land_price, struc_price, p_land_price, property_value2, area1[0], area2[0], area1[1], area2[1], chung, sale_number, agent_name, agent_phone, jungong_date, memo])
        except:
            continue
columns = ['지하철역', '매매가', '대출', '보증금', '월세', '4.5%', '확인유무', '토지가격', '건물가격', '평당토지', '4%', '대지면적', '건축면적', '연면적', '전용면적', '층수', '매물번호', '중개인', '전화번호', '사용승인일', '특징']
real_df = pd.DataFrame(real_list, columns=columns)
real_df.head()
real_df.info()
real_df.to_excel('./부산지역_부동산매물전체_naver_list ' + curr_date + '.xlsx', index=True)