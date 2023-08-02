from selenium import webdriver
from selenium.webdriver.common.by import By
#from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import openpyxl

lists = []
lists1 = []
real_list = []
i = 0
# 네이버 부동산 실행
url = 'https://map.naver.com/'
curr_date = datetime.strftime(datetime.now(), "%y_%m_%d")
browser = webdriver.Chrome(ChromeDriverManager().install())
browser.get(url)
url = input('검색지역의 url을 복사해 넣으세요 : ')
browser.get(url)
search_area = input("검색지역명을 입력하세요 : ")
search_business = input("검색할 사업종류를 입력하세요 : ")
time.sleep(1)

# 조회할 지역의 자료를 페이지다운으로 스크롤 하고 지역명 입력
browser.switch_to.frame('searchIframe')

wait = WebDriverWait(browser, 20)
element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#_pcmap_list_scroll_container')))

verical_ordinate = 100
for i in range(0, 5):
   browser.execute_script("arguments[0].scrollTop = arguments[1]", element, verical_ordinate)
   verical_ordinate += 1000
   time.sleep(1)

html = browser.page_source
soup = BeautifulSoup(html, 'html.parser')
lists = soup.select('#_pcmap_list_scroll_container > ul > li')
i = 1
browser.switch_to.default_content()
time.sleep(1)
for list in lists:
    browser.switch_to.frame('searchIframe')
    time.sleep(1)
    body = browser.find_element(By.CSS_SELECTOR, '#_pcmap_list_scroll_container > ul > li:nth-child(' + str(i) + ') > div._3hn9q > a > div.O9Z-o > div')
    body.click()
    time.sleep(1)
    i += 1
# 매물 하나씩 선택하기
    browser.switch_to.default_content()
    browser.switch_to.frame('entryIframe')
    time.sleep(1.5)
    title = browser.find_element(By.CSS_SELECTOR, '#_title > span._3XamX').text
    print(title)
    # address = wait.until(browser.find_element(By.CSS_SELECTOR, '#app-root > div > div > div.place_detail_wrapper > div:nth-child(4) > div > div:nth-child(2) > div > ul > li._1M_Iz._1aj6- > div > a > span._2yqUQ'))
    # print(address.text)
    bodys = browser.find_elements(By.CSS_SELECTOR, '#app-root > div > div > div.place_detail_wrapper > div.place_fixed_maintab > div > div > div > div > a > span._3aXen')
# 리뷰 메뉴 선택하기
    count = len(bodys)
    for body1 in bodys:
        if body1.text == "리뷰":
            body1.click()
            break
    time.sleep(1.5)
    if body1.text == "리뷰":
        html = browser.page_source
        soup1 = BeautifulSoup(html, 'html.parser')
        lists1 = soup1.find_all('li', '_3FaRE')
        if len(lists1) == 0:
            browser.switch_to.default_content()
            continue
        reviews = []
        for list1 in lists1[:5]:
            memo = list1.find('span', '_1lntw')
            if memo == None:
                browser.switch_to.default_content()
                break
            else:
                memo = memo.text.strip('"')
            number = list1.find('span', 'Nqp-s').text.strip("이 키워드를 선택한 인원")
            reviews.append([memo, number])
        browser.switch_to.default_content()
        if memo == None:
            continue
        else:
            real_list.append([title, reviews[0][0], reviews[0][1], reviews[1][0], reviews[1][1], reviews[2][0], reviews[2][1], reviews[3][0], reviews[3][1], reviews[4][0], reviews[4][1]])
    else:
        browser.switch_to.default_content()
        continue
columns = ['카페이름', '방문이유 1', '인원', '방문이유 2', '인원', '방문이유 3', '인원', '방문이유 4', '인원', '방문이유 5', '인원']
real_df = pd.DataFrame(real_list, columns=columns)
real_df.head()
real_df.info()
real_df.to_excel('./' + search_area + '지역 ' + search_business + ' ' + curr_date + '.xlsx', index=True)


soup = BeautifulSoup(html, 'html.parser')
lists = soup.select('#_pcmap_list_scroll_container > ul > li')
i = 1
browser.switch_to.default_content()
for list in lists:
    browser.switch_to.frame('searchIframe')
    body = browser.find_element(By.CSS_SELECTOR, '#_pcmap_list_scroll_container > ul > li:nth-child(' + str(i) + ') > div._3hn9q > a > div.O9Z-o > div')
    body.click()
    time.sleep(1)
    i += 1
# 매물 하나씩 선택하기
    browser.switch_to.default_content()
    browser.switch_to.frame('entryIframe')
    time.sleep(1.5)
    title = browser.find_element(By.CSS_SELECTOR, '#_title > span._3XamX').text
    print(title)
    # address = wait.until(browser.find_element(By.CSS_SELECTOR, '#app-root > div > div > div.place_detail_wrapper > div:nth-child(4) > div > div:nth-child(2) > div > ul > li._1M_Iz._1aj6- > div > a > span._2yqUQ'))
    # print(address.text)
    bodys = browser.find_elements(By.CSS_SELECTOR, '#app-root > div > div > div.place_detail_wrapper > div.place_fixed_maintab > div > div > div > div > a > span._3aXen')
# 리뷰 메뉴 선택하기
    count = len(bodys)
    for body1 in bodys:
        if body1.text == "리뷰":
            body1.click()
            break
    time.sleep(1.5)
    if body1.text == "리뷰":
        html = browser.page_source
        soup1 = BeautifulSoup(html, 'html.parser')
        lists1 = soup1.find_all('li', '_3FaRE')
        if len(lists1) == 0:
            browser.switch_to.default_content()
            continue
        reviews = []
        for list1 in lists1[:5]:
            memo = list1.find('span', '_1lntw')
            if memo == None:
                browser.switch_to.default_content()
                break
            else:
                memo = memo.text.strip('"')
            number = list1.find('span', 'Nqp-s').text.strip("이 키워드를 선택한 인원")
            reviews.append([memo, number])
        browser.switch_to.default_content()
        if memo == None:
            continue
        else:
            real_list.append([title, reviews[0][0], reviews[0][1], reviews[1][0], reviews[1][1], reviews[2][0], reviews[2][1], reviews[3][0], reviews[3][1], reviews[4][0], reviews[4][1]])
    else:
        browser.switch_to.default_content()
        continue
columns = ['카페이름', '방문이유 1', '인원', '방문이유 2', '인원', '방문이유 3', '인원', '방문이유 4', '인원', '방문이유 5', '인원']
real_df = pd.DataFrame(real_list, columns=columns)
real_df.head()
real_df.info()
real_df.to_excel('./' + search_area + '지역 ' + search_business + ' review ' + curr_date + '.xlsx', index=True)

