import OpenDartReader
import requests                     #post로 dcmno 요청하기 위함
import re
from bs4 import BeautifulSoup       #html로 쉽게 파싱하기 위함.
from openpyxl import Workbook       #엑셀로 출력하는 도구 openpyxl을 pip에서 설치하셔야 합니다.

#변수 설명
#pdf : pdf 경로 url
#data[] : 새창 보고서 내 문서 별 변수 (rcpno, dcmno, eleid, offset등...) 이것으로 문서를 왔다갔다하는 url을 만든다.
#rcp_no[] : rcp_no를 모아둔 배열
#name_reports[] : 보고서 이름을 모아둔 배열
#name_company[] : 회사명을 모아둔 배열
#corp_code : 보고서 별로 comp_code의 위치가 다르다. 현재 미완성. 돌려보면서 경우의 수마다 새로 생성 예정

#홍차장님 증정
api_key='8fd02dac927493a7161e13a34e78062b49197e59'
#dart = OpenDartReader(api_key)

date='2024.02.08'
url="https://dart.fss.or.kr/dsac001/mainF.do?selectDate="+date+"&sort=&series=&mdayCnt=0"
#날짜가 잘못되면 오류를 뿜기 때문에 아래 구문은 추후 try-catch로 감싸줘야한다.
page=requests.get(url)
html=page.text
soup = BeautifulSoup(html, 'html.parser')

#페이지 내 날짜를 구한다.
path='#listContents > div.tbTitle > h4'
target=soup.select_one(path)
pdate=' '.join(target.get_text().split()[2:])[1:-1]

#다음 페이지의 갯수를 구한다. 가져올 보고서 대상 결정
path='#listContents > div.psWrap > div.pageSkip > ul'
li=target.find_all('li')
if len(li)>1:
    for i in range(1,len(li)):
        #모든 페이지의 값을 가져온다. post를 날리면 다음 페이지의 html을 가져온다.
        frm_page = soup.find('form', attrs={'name': 'searchForm'})
        url = 'https://dart.fss.or.kr/dsac001/mainF.do'
        data = {
            'currentPage': str(i),
            'maxResults': '',
            'maxLinks': '',
            'sort': '',
            'series': '',
            'pageGrouping': 'F',
            'mdayCnt': '0',
            'selectDate': date,
            'textCrpCik': ''
        }
        response = requests.post(url, data=data)


#현재 페이지의 테이블 내 보고서 리스트를 구함
path='#listContents > div.tbListInner > table > tbody'
target=soup.select_one(path)
reports=[]
name_company=[]
if target :
    tr_tags=target.find_all('tr')
    no_tr=len(tr_tags)
    for i in range(no_tr):
        reports.append(str(tr_tags[i].find_all('a')[1]).split())
        name_company.append(str(tr_tags[i].find_all('a')[0]).split()[-2])
else :
    print("공시가 없습니다.")  #물론 그럴리는 없음

rcp_no=[]
name_reports=[]
for i in reports:
    rcp_no.append(i[1].split('=')[2].strip().replace('\"',''))
    name_reports.append(i[6].split('=')[1])

#data 정보를 읽어온다.
url='https://dart.fss.or.kr/dsaf001/main.do?rcpNo='+rcp_no[0]

#"일괄신고서"에서 가져오는 경우

#pdf 다운로드 url
pdf_url='https://dart.fss.or.kr/pdf/download/main.do?rcp_no='+data[0]['rcpNo']+'&dcm_no='+data[0]['dcmNo']
#pdf 목록 tbody selector
target='body > div > div.cont > div > div > table > tbody'
soup = BeautifulSoup(requests.get(pdf_url).text, 'html.parser')
pdf=soup.select_one(target)
pdf='http://dart.fss.or.kr'+pdf.select_one('a.btnFile').get('href') #pdf 경로 완성

page=requests.get(url)
html=page.text
var=re.findall(r'node(.*);',html)
arr=[]
for i in var:
    temp=[]
    try:
        temp.append(re.findall(r'\'(.*)\'',i)[0])               #col
        temp.append(i.split('=')[1].strip().replace('\"',''))   #value
        arr.append(temp)
    except:
        continue

data=[]
#var을 dict변수로 생성한다.
cur={}

for i in arr:
    if i[0]=='text':
        if cur:
            data.append(cur)
            cur={}
    cur[i[0]]=i[1]

if cur:
    data.append(cur)

#명칭이 있는 페이지에서 펀드번호를 받아온다. "집합투자기구의 명칭" => 여기서 펀드코드를 가져온다. => temp
url='https://dart.fss.or.kr/report/viewer.do?rcpNo=20240115000363&dcmNo=9576824&eleId=2&offset=4599&length=670&dtd=dart3.xsd'
page=requests.get(url)
soup = BeautifulSoup(page.text, 'html.parser')
temp=soup.select_one('body > table > tbody > tr:nth-child(2) > td:nth-child(1)')
print(temp.get_text())
temp=soup.select_one('body > table > tbody > tr:nth-child(2) > td:nth-child(2)')
print(temp.get_text())

#"일괄신고서"에서 가져오는 경우

#pdf 다운로드 url
pdf_url='https://dart.fss.or.kr/pdf/download/main.do?rcp_no='+data[0]['rcpNo']+'&dcm_no='+data[0]['dcmNo']
#pdf 목록 tbody selector
target='body > div > div.cont > div > div > table > tbody'
soup = BeautifulSoup(requests.get(pdf_url).text, 'html.parser')
pdf=soup.select_one(target)
pdf='http://dart.fss.or.kr'+pdf.select_one('a.btnFile').get('href') #pdf 경로 완성


