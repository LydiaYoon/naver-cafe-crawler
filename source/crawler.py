# -*- coding:utf-8 -*-

import getpass
import openpyxl
import os
import random
import time
from bs4 import BeautifulSoup
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from urllib.request import urlopen
from urllib.error import URLError, HTTPError

LOADING_TIMEOUT = 3

# 법적 책임 동의
print("이 프로그램 사용으로 인한 모든 법적 책임은 전적으로 사용자에게 있습니다.\n동의하십니까? (Y/N)")
confirm = input().strip() 
if not confirm == 'Y' :
    exit()
    
print('\n==================================================')
# 입력 파일명
print("\n입력 파일명을 입력해주세요.")
importFileName = input("file name: ").strip()

# 출력 파일명
print("\n출력 파일명을 입력해주세요.")
exportFileName = input("export name: ").strip()

# 네이버 계정
print("\n계정 정보를 입력해주세요.")
id = input("naver id: ").strip()
pw = getpass.getpass(prompt="naver pw: ").strip()


print('\n==================================================\n')
print('입력 파일 확인중')
try:
    # 엑셀 파일 읽기
    fileUrl = os.getcwd() + "/" + importFileName + ".xlsx"
    
    # data_only=True로 해줘야 수식이 아닌 값으로 받아온다.
    book = openpyxl.load_workbook(fileUrl, data_only=True)
    
    # 활성화된 시트 추출하기
    sheet = book.active
except FileNotFoundError as e:
    print("입력 파일명이 잘못되었습니다.")
    print(e)
    exit()
print('입력 파일 확인 완료')

    
print('\n==================================================\n')
print('네이버 로그인')

# 크롬 드라이버 연결
driver = webdriver.Chrome(os.getcwd() + '/driver/chromedriver.exe')

# 네이버 로그인 페이지 접속 
driver.get('https://nid.naver.com/nidlogin.login')
driver.implicitly_wait(LOADING_TIMEOUT)  # 로딩이 끝날 때 까지 대기 
    
# driver.maximize_window() # 브라우저 화면 최대화
 
# 네이버 로그인
driver.execute_script("document.getElementsByName('id')[0].value=\'" + id + "\'")
time.sleep(random.randrange(1, 3))
driver.execute_script("document.getElementsByName('pw')[0].value=\'" + pw + "\'")
time.sleep(random.randrange(1, 3))
driver.find_element_by_class_name('btn_global').click()
time.sleep(random.randrange(1, 3))

print('네이버 로그인 확인중')
html = driver.page_source  # 현재 페이지의 주소를 반환 
soup = BeautifulSoup(html, 'html.parser')
loginError = soup.find_all(class_='error', attrs={"aria-live": "assertive"})
if loginError:
    print("가입하지 않은 아이디이거나, 잘못된 비밀번호입니다.")
    exit()
print('네이버 로그인 완료')
    
print('\n==================================================\n')
print('데이터 크롤링')
print('입력 파일의 행 개수: ', sheet.max_row)

for i in range(1, sheet.max_row):  # 4행부터 229번째 행 까지
    # 읽기
    print()
    print(i, "번째 행")
    postUrl = sheet['G' + str(i)].value  # G열 (URL)
    print("게시물 주소:", postUrl)
    
    if not postUrl:
        continue
    
    try:
        res = urlopen(postUrl)
        # print(res.status)
    except ValueError as e:
        print("올바르지 않은 주소 형식입니다.")
        print(e)
        continue
    except HTTPError as e:
        err = e.read()
        code = e.getcode()
        print(code)  # # 404
    
    # 네이버 카페 페이지 들어가기
    driver.get(postUrl)
    driver.implicitly_wait(LOADING_TIMEOUT)  # 로딩이 끝날 때 까지 대기
     
    try:  # 경고창이 있을 경우
        WebDriverWait(driver, 1).until(EC.alert_is_present(),
                                       'Timed out waiting for PA creation ' + 
                                       'confirmation popup to appear.')
     
        alert = driver.switch_to.alert
        print(alert.text)  # 삭제되었거나 없는 게시글입니다.
        alert.accept()
        sheet['I' + str(i)] = '삭제된 게시물'  # I열 (댓글수)
        sheet['H' + str(i)] = '삭제된 게시물'  # H열 (조회수)
         
    except TimeoutException:  # 경고창이 없을 경우
        html = driver.page_source  # 현재 페이지의 주소를 반환 
        soup = BeautifulSoup(html, 'html.parser')
              
        # 카페 포스트 본문을 보여주는 iframe 주소를 찾는다
        iframes = soup.find_all('iframe', id="cafe_main")
              
        # for iframe in iframes:
        #    print(iframe.get('name'))
              
        driver.switch_to_default_content  # 상위 프레임으로 전환
        driver.switch_to.frame('cafe_main')  # cafe_main 프레임으로 전환
              
        html = driver.page_source  # 현재 페이지의 주소를 반환 
        soup = BeautifulSoup(html, 'html.parser')
        
        # 댓글수
        replyCount = ''
        replyEl = soup.find_all(class_='ArticleTool')
        if replyEl and replyEl[0].find_all('strong', class_='num'):
            replyCount = int(replyEl[0].find_all('strong', class_='num')[0].text.strip().replace(',', ''))
        else:
            replyCount = '댓글수 찾기 불가'
        
        sheet['I' + str(i)] = replyCount
         
        # 조회수
        viewCount = ''
        viewEl = soup.find_all(class_='article_info')
        if viewEl and viewEl[0].find_all('span', class_='count'):
            viewCount = int(viewEl[0].find_all('span', class_='count')[0].text.strip()[3:].replace(',', ''))
        else:
            viewCount = '조회수 찾기 불가'
        sheet['H' + str(i)] = viewCount
        
    print("댓글수:", sheet['I' + str(i)].value)
    print("조회수:", sheet['H' + str(i)].value)

time.sleep(random.randrange(1, 3))
driver.close()

print('데이터 크롤링 완료')
         
print('\n==================================================\n')

print('출력 파일 저장')
# 엑셀 파일 저장하기
currentTime = datetime.today().strftime("%y%m%d_%H%M%S")
fileUrl = exportFileName + "_" + currentTime + ".xlsx"
book.save(fileUrl)
print('출력 파일 저장 완료')
print("결과가 \"" + fileUrl + "\"으로 저장되었습니다.")
    
