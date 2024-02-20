import OpenDartReader
import requests                     #post로 dcmno 요청하기 위함
import re
from bs4 import BeautifulSoup       #html로 쉽게 파싱하기 위함.
from openpyxl import Workbook       #엑셀로 출력하는 도구 openpyxl을 pip에서 설치하셔야 합니다.
#requests로 그냥하면 막힌다. tor로 프록시를 깔려했으나 페이지가 보안에 막힘 ㅠ 그냥 selenium으로 긁자.
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from datetime import datetime
import time

def list_to_dict(lst):
    result = {}
    for item in lst:
        key, value = item
        result[key] = value
    return result

def is_element_present(driver, by, value):
    try:
        driver.find_element(by=by, value=value)
        return True
    except Exception:
        return False

#크롤링 건마다 쉬는 속도
MAX_SLEEP_TIME=3000
#펀드공시 페이지 정보
reports=[]          

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
}

date='2024.02.19'
#date=datetime.today().strftime('%Y.%m.%d')
url="https://dart.fss.or.kr/dsac001/mainF.do?selectDate="+date+"&sort=&series=&mdayCnt=0"

#디버깅 크롬으로 코드를 바꿈. 크롬을 실행시키고 

options = webdriver.ChromeOptions()
#options.add_argument('--headless')
options.add_argument('--disable-gpu')
#options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

service = Service('c:\chromedriver.exe')
driver = webdriver.Chrome('c:\chromedriver.exe', options=options)
driver.get(url)
#페이지가 뜨길 기다린다.
time.sleep(1)

#페이지 내 날짜를 구한다.
path='#listContents > div.tbTitle > h4'
target=driver.find_element(By.CSS_SELECTOR, path)
pdate=' '.join(target.text.split()[2:])[1:-1]

#다음 페이지의 갯수를 구한다. 가져올 보고서 대상 결정
path='#listContents > div.psWrap > div.pageSkip > ul'
btn_next=driver.find_element(By.CSS_SELECTOR, path)
li=btn_next.find_elements_by_tag_name('li')   #하단 다음 버튼 갯수
pages=[]
pos=0
main_page = driver.current_window_handle
for i in range(len(li)):
    #펀드 공시 목록을 하나하나 클릭하면서 값을 찾는다. dcmno고 자시고 필요가 없다.
    path='#listContents > div.tbListInner > table'
    target=driver.find_element(By.CSS_SELECTOR, path)
    report=target.find_elements_by_tag_name('tr') 
    del report[0]  #시간, 공시대상회사 등의 legend 행을 날린다.
    for j in range(len(report)):
        temp=[]
        #하나하나 뺑글뺑글 돌자.
        line_data=report[j].find_elements_by_tag_name('td')
        path_corpcode='#listContents > div.tbListInner > table > tbody > tr:nth-child('+str(i+1)+') > td:nth-child(2) > span > a'
        corp_code=driver.find_element(By.CSS_SELECTOR, path_corpcode).get_attribute('href').split("\'")[1]
        temp.append(corp_code)                                                  #cort_code
        temp.append(line_data[1].text.split()[1])                               #공시회사명
        print(str(i)+' : '+temp[-1])
        temp.append(line_data[2].text)                                          #보고서명
        btn_report=line_data[2].find_element_by_tag_name('a')
        temp.append(btn_report.get_attribute('href').split('=')[1])             #rcp_no
        temp.append(line_data[0].text)                                          #시간
        temp.append(line_data[3].text)                                          #제출인
        temp.append(line_data[4].text)                                          #접수일자
        temp.append('https://dart.fss.or.kr/dsaf001/main.do?rcpNo='+temp[3])    #report_link
        temp.append('')                                                         #펀드코드 (일단 비워둔다)
        btn_report.click()
        #새 창이 뜨길 기다립니다..
        time.sleep(1)
        driver.switch_to.window(driver.window_handles[1])                       #창 바꿈         
        report_page = driver.current_window_handle
        path_down='body > div.wrapper > div > div.header > div.bottom > div.rightWrap > button.btnDown'
        try:
            #펀드코드를 가져온다.
            #좌측 문서 목차는 iframe로 화면을 구성해서 페이지 url을 만들어 접속을 시도해야한다.... 쌍팔년도인줄
            #eleid, offset, length를 받아와야한다.
            path='#listTree > ul'
            target=driver.find_element(By.CSS_SELECTOR, path)
            #node의 값을 읽어 offset, eleid, length의 값을 가져온다.
            nodes=re.findall(r'node1(.*);', driver.page_source)
            data=[]
            for k in nodes:
                k=re.sub(r'\s+|\[|\]|\'|\"', '', k) #각 줄에서 공백, [ ] '을 제거
                data.append(k.split('='))
            nodes=[]
            start=0
            end=0
            for k in range(len(data)):
                if data[k][0]=='':
                    start=k
                elif data[k][0]==')':
                    end=k
                    nodes.append(list_to_dict(data[start+1:end]))
            #nodes에 문서 목차 내 변수를 다 받아왔다.
            urls=[]
            for k in nodes:
                temp_url='https://dart.fss.or.kr/report/viewer.do?rcpNo='+k['rcpNo']+'&dcmNo='+k['dcmNo']+'&eleId='+k['eleId']+'&offset='+k['offset']+'&length='+k['length']+'&dtd='+k['dtd']
                urls.append(temp_url)
            #urls(문서목차)를 돌면서 펀드코드를 찾는다.
            for k in range(len(urls)):
                time.sleep(1)   #혹시 모를 크롤링 짤 방지
                driver.execute_script("window.open('about:blank', '"+nodes[k]['text']+"');")  #새창을 띄운다.
                driver.switch_to.window(driver.window_handles[-1])
                driver.get(urls[k])
                text=driver.find_element(By.CSS_SELECTOR, 'body').text
                if '펀드코드' in text:
                    #몇가지 펀드코드 위치 경우의 수 에서 찾아보고 없으면 포기
                    path1='body > table:nth-child(6) > tbody > tr:nth-child(1) > td:nth-child(4)'   #투자설명서
                    if is_element_present(driver,By.CSS_SELECTOR, path1):
                        target=driver.find_element(By.CSS_SELECTOR, path1)
                        temp[8]=target.text
                        driver.close()
                        driver.switch_to.window(driver.window_handles[-1])
                        break
                driver.close()
                driver.switch_to.window(driver.window_handles[-1])
            #다운로드 창을 켠다.
            if is_element_present(driver,By.CSS_SELECTOR, path_down):
                btn_down=driver.find_element(By.CSS_SELECTOR, path_down)
                temp.append(btn_down.get_attribute('onclick').split(',')[1][2:-3])  #dcm_no
                btn_down.click()
                #새 창이 뜨길 기다립니다..
                time.sleep(0.5)
                driver.switch_to.window(driver.window_handles[2])
                path_table='body > div > div.cont > div > div > table'
                table=driver.find_element(By.CSS_SELECTOR, path_table)
                list_pdf=table.find_elements_by_tag_name('tr')
                del list_pdf[0]
                for tr in list_pdf:
                    td=tr.find_elements_by_tag_name('td')
                    temp.append(td[0].text)                                                 #pdf명
                    temp.append(td[1].find_element_by_tag_name('a').get_attribute('href'))  #pdf url
                #다운로드 창을 끈다     
                driver.close()
                driver.switch_to.window(report_page)
            #메인화면으로 돌아온다
            driver.close()
            driver.switch_to.window(main_page)                
        except Exception as e:
            temp.append('')         #dcm_no
            temp.append('')         #pdf명
            temp.append('')         #pdf url
            #메인화면으로 돌아온다
            driver.close()
            driver.switch_to.window(main_page)
            print(e)

        #공시 정보를 담는다.
        reports.append(temp)            
        

    #다음페이지로
    pos+=1
    if(pos!=len(li)):
        #매 페이지마다 li를 새로 만들어 줘야한다.. 왠지모르게 있던거 쓰면 안된다. 어디서 건드린듯.
        path='#listContents > div.psWrap > div.pageSkip > ul'
        btn_next=driver.find_element(By.CSS_SELECTOR, path)
        li=btn_next.find_elements_by_tag_name('li')   #하단 다음 버튼 갯수
        li[pos].find_element_by_tag_name('a').click()
        time.sleep(0.5)

driver.close()

write_wb = Workbook()
write_ws = write_wb.create_sheet('펀드공시'+'_'+date)
#Sheet1에다 입력
write_ws = write_wb.active
write_ws['A1'] = 'corp_code'    #추가
write_ws['B1'] = 'corp_name'
write_ws['C1'] = 'report_nm'
write_ws['D1'] = 'rcept_no'
write_ws['E1'] = 'upld_tm'
write_ws['F1'] = 'flr_nm'
write_ws['G1'] = 'rcept_dt'
write_ws['H1'] = 'rcept_link'
write_ws['I1'] = 'fund_code'
write_ws['J1'] = 'dcm_no'
write_ws['K1'] = 'file_nm1'
write_ws['L1'] = 'file_link1'
write_ws['M1'] = 'file_nm2'
write_ws['N1'] = 'file_link2'
write_ws['O1'] = 'file_nm3'
write_ws['P1'] = 'file_link3'

for i in reports:
    write_ws.append(i)
#행 단위로 추가
write_wb.save("C:/펀드기업공시_["+date+"].xlsx")
