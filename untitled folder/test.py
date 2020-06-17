# -*- coding:utf-8 -*-
import csv
from collections import OrderedDict
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup

data = []
with open('/Users/narae/Downloads/test-csv.csv', 'r', encoding='utf-8-sig') as csv_file:
    reader = csv.DictReader(csv_file, skipinitialspace=True)
     
    temp = list(reader)
 
    for row in temp:
        if (row.get('URL') != ""):
            r = OrderedDict()
            r['순번'] = row.get('순번')
            r['업로드일자'] = row.get('업로드일자')
            r['항목'] = row.get('항목')
            r['카페명'] = row.get('카페명')
            r['회원수'] = row.get('회원수')
            r['URL'] = row.get('URL')
            r['조회수'] = row.get('조회수')
            r['댓글수'] = row.get('댓글수')
            
            data.append(r)
 
for row in data:
        print(row)
         
# 크롬 드라이버 연결
driver = webdriver.Chrome('/usr/local/bin/chromedriver')
    
# 네이버 로그인 페이지 접속 
driver.get('https://nid.naver.com/nidlogin.login')
# driver.get('https://cafe.naver.com/sqlpd/9910')
driver.implicitly_wait(5) # 로딩이 끝날 때 까지 대기 
    
# driver.maximize_window() # 브라우저 화면 최대화
 
id = 'f_ake'
pw = 'a9024096'
driver.execute_script("document.getElementsByName('id')[0].value=\'" + id + "\'")
time.sleep(1)
driver.execute_script("document.getElementsByName('pw')[0].value=\'" + pw + "\'")
time.sleep(1)
driver.find_element_by_class_name('btn_global').click()
time.sleep(1)
  
for row in data:
    post_url = row.get('URL')
    print(post_url)
          
    # 네이버 카페 페이지 들어가기
    driver.get(post_url)
    driver.implicitly_wait(5) # 로딩이 끝날 때 까지 대기
     
    try: # 경고창이 있을 경우
        WebDriverWait(driver, 1).until(EC.alert_is_present(),
                                       'Timed out waiting for PA creation ' +
                                       'confirmation popup to appear.')
     
        alert = driver.switch_to.alert
        alert.accept()
        row['댓글수'] = '삭제된 게시물'
        row['조회수'] = '삭제된 게시물'
         
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
         
        reply = soup.find_all(class_='_totalCnt')[0].text.strip()[3:]
         
        if soup.find_all(class_='b m-tcol-c reply'):
            views = soup.find_all('span', class_='b m-tcol-c reply')[1].text.strip().replace(',','')
        else:
            views = '조회수 찾기 불가'
        # views = soup.find_all('span', class_='b m-tcol-c reply')[1].text.strip()
        print(reply)
        print(views)
        row['댓글수'] = reply
        row['조회수'] = views
         
time.sleep(5)
driver.close()
 
# csv 파일로 결과 출력 
with open('/Users/narae/Downloads/test-res.csv', 'w') as csv_file:
    fieldnames = ['순번', '업로드일자', '항목', '카페명','회원수', 'URL', '조회수', '댓글수']
    writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
    writer.writeheader()
    for row in data:
        writer.writerow(row)
                
                
