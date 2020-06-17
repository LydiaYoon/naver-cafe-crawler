# -*- coding:utf-8 -*-

import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup

# 크롬 드라이버 연결
driver = webdriver.Chrome('/usr/local/bin/chromedriver')
    
# 네이버 로그인 페이지 접속 
driver.get('https://nid.naver.com/nidlogin.login')
driver.implicitly_wait(3) # 로딩이 끝날 때 까지 대기 
    
# driver.maximize_window() # 브라우저 화면 최대화
 
# 네이버 로그인
id = 'user_id' # 사용자 아이디
pw = 'user_password' # 비밀번호
driver.execute_script("document.getElementsByName('id')[0].value=\'" + id + "\'")
time.sleep(1)
driver.execute_script("document.getElementsByName('pw')[0].value=\'" + pw + "\'")
time.sleep(1)
driver.find_element_by_class_name('btn_global').click()
time.sleep(1)

# 엑셀 파일 읽기
file_url = "/Users/narae/Downloads/test.xlsx"

# data_only=Ture로 해줘야 수식이 아닌 값으로 받아온다.
book = openpyxl.load_workbook(file_url, data_only=True)

# 활성화된 시트 추출하기
sheet = book.active

for i in range(4, 230): # 4행부터 229번째 행 까지
    # 읽기
    post_url = sheet['G' + str(i)].value # G열 (URL)
    
    # 네이버 카페 페이지 들어가기
    driver.get(post_url)
    driver.implicitly_wait(3) # 로딩이 끝날 때 까지 대기
     
    try: # 경고창이 있을 경우
        WebDriverWait(driver, 1).until(EC.alert_is_present(),
                                       'Timed out waiting for PA creation ' +
                                       'confirmation popup to appear.')
     
        alert = driver.switch_to.alert
        alert.accept()
        sheet['I' + str(i)] = '삭제된 게시물' # I열 (댓글수)
        sheet['H' + str(i)] = '삭제된 게시물' # H열 (조회수)
         
    except TimeoutException: # 경고창이 없을 경우
     
        html = driver.page_source # 현재 페이지의 주소를 반환 
        soup = BeautifulSoup(html, 'html.parser')
              
        # 카페 포스트 본문을 보여주는 iframe 주소를 찾는다
        iframes = soup.find_all('iframe', id="cafe_main")
              
        for iframe in iframes:
            print(iframe.get('name'))
              
        driver.switch_to_default_content # 상위 프레임으로 전환
        driver.switch_to.frame('cafe_main') # cafe_main 프레임으로 전환
              
        html = driver.page_source # 현재 페이지의 주소를 반환 
        soup = BeautifulSoup(html, 'html.parser')
         
        # 댓글수와 조회수를 찾는다     
        reply_sort = soup.find_all('div', class_='fl reply_sort')
        
        # 댓글수
        sheet['I' + str(i)] = int(soup.find_all(class_='_totalCnt')[0].text.strip()[3:].replace(',',''))
         
        # 조회수
        if soup.find_all(class_='b m-tcol-c reply'):
            sheet['H' + str(i)] = int(soup.find_all('span', class_='b m-tcol-c reply')[1].text.strip().replace(',',''))
        else:
            sheet['H' + str(i)] = '조회수 찾기 불가'
            
    print("게시물 주소:", post_url)
    print("댓글수:", sheet['I' + str(i)].value)
    print("조회수:", sheet['H' + str(i)].value)

time.sleep(3)
driver.close()
         
# 엑셀 파일 저장하기
file_url = "/Users/narae/Downloads/test-res.xlsx"
book.save(file_url)
print("OK")
