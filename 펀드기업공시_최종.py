import OpenDartReader
import requests                     #post로 dcmno 요청하기 위함
import re
from bs4 import BeautifulSoup       #html로 쉽게 파싱하기 위함.
from openpyxl import Workbook       #엑셀로 출력하는 도구 openpyxl을 pip에서 설치하셔야 합니다.

#크롤링 건마다 쉬는 속도
MAX_SLEEP_TIME=3000
#펀드공시 페이지 정보
rcp_no=[]           #rcp_no
dcm_no=[]           #dcm_no
reports=[]          #페이지별 html > 여기서 리포트명과 rcpno를 파싱한다.
name_reports=[]     #리포트명
name_company=[]     #회사명
#공시별 pdf 관련 정보 (새창)
fund_code=[]        #펀드 코드 (종류별 처리 필요)
pdf_url=[]          #한 리포트에 pdf가 여러개인 경우가 있어서 이차원배열로 사용
doc_data=[]         #리포트 관련 정보

#홍차장님 증정
api_key='8fd02dac927493a7161e13a34e78062b49197e59'

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
}

date='2024.02.08'
url="https://dart.fss.or.kr/dsac001/mainF.do?selectDate="+date+"&sort=&series=&mdayCnt=0"
#날짜가 잘못되면 오류를 뿜기 때문에 아래 구문은 추후 try-catch로 감싸줘야한다.
page=requests.get(url, headers=headers)
html=page.text
soup = BeautifulSoup(html, 'html.parser')   #lxml이 좀 더 빠르나 우선 html로 구현

#페이지 내 날짜를 구한다.
path='#listContents > div.tbTitle > h4'
target=soup.select_one(path)
pdate=' '.join(target.get_text().split()[2:])[1:-1]

#다음 페이지의 갯수를 구한다. 가져올 보고서 대상 결정
path='#listContents > div.psWrap > div.pageSkip > ul'
target=soup.select_one(path)
li=target.find_all('li')
pages=[]
if len(li)>1:
    for i in range(len(li)):
        #모든 페이지의 값을 가져온다. post를 날리면 다음 페이지의 html을 가져온다.
        frm_page = soup.find('form', attrs={'name': 'searchForm'})
        url = 'https://dart.fss.or.kr/dsac001/mainF.do'
        data = {
            'currentPage': str(i+1),
            'maxResults': '',
            'maxLinks': '',
            'sort': '',
            'series': '',
            'pageGrouping': 'F',
            'mdayCnt': '0',
            'selectDate': date,
            'textCrpCik': ''
        }
        response = requests.post(url, data=data, headers=headers)
        pages.append(response)
        html=response.text
        soup = BeautifulSoup(html, 'html.parser')

        #현재 페이지의 테이블 내 보고서 리스트를 구함
        path='#listContents > div.tbListInner > table > tbody'
        target=soup.select_one(path)
        if target :
            tr_tags=target.find_all('tr')
            no_tr=len(tr_tags)
            for i in range(no_tr):
                reports.append(str(tr_tags[i].find_all('a')[1]).split())
                name_company.append(str(tr_tags[i].find_all('a')[0]).split()[-2])
        else :
            print("공시가 없습니다.")  #물론 그럴리는 없음

#reports 정보를 다 받아왔다. 페이지 내의 정보를 가져오자.
for i in reports:
    rcp_no.append(i[1].split('=')[2].strip().replace('\"',''))
    end=i.index('공시뷰어')
    name_reports.append(''.join(i[6:end]).split('=')[1])

#각 리포트별로 펀드코드와 pdf 정보를 가져온다. 새창을 연다.
for i in range(len(rcp_no)):
    #조금만 쉬자.. 막는다..
    rand_value = randint(1, MAX_SLEEP_TIME)
    time.sleep(rand_value)
    
    url='https://dart.fss.or.kr/dsaf001/main.do?rcpNo='+rcp_no[i]
    page=requests.get(url, headers=headers)
    html=page.text
    soup = BeautifulSoup(html, 'html.parser')

    #js변수에서 문서 목록과 dcmno를 가져온다.
    var=re.findall(r'node(.*)',html)
    arr=[]
    for j in var:
        temp=[]
        try:
            temp.append(re.findall(r'\'(.*)\'',j)[0])               #col
            temp.append(j.split('=')[1].strip().replace('\"',''))   #value
            arr.append(temp)
        except:
            continue
        
    #var을 dict변수로 생성한다.
    data=[]
    cur={}
    for j in arr:
        if j[0]=='text':
            if cur:
                data.append(cur)
                cur={}
        cur[j[0]]=j[1]
    data.append(cur)    #마지막 하나 더 넣음

    #dcm_no 입력
    try:
        dcm_no.append(data[0]['dcmNo'][:-1])
    except:
        #효력발생안내만 있고 아무것도 없는 화면들이 있다. 그냥 넘어가자
        dcm_no.append('')
        pdf_url.append([])
        continue

    #여기서 좌측 문서를 확인하여 fund_code를 구한다. (구현예정)

    #pdf 다운로드 url
    pdf=[]
    url_pdf='https://dart.fss.or.kr/pdf/download/main.do?rcp_no='+rcp_no[i]+'&dcm_no='+data[0]['dcmNo'][:-1]
    response = requests.get(url_pdf, headers=headers)
    html=response.text
    soup=BeautifulSoup(html, 'html.parser')

    target='body > div > div.cont > div > div > table > tbody'
    table=soup.select_one(target)
    tr=table.find_all('tr')
    for j in tr:
        pdf.append(j.find_all('td')[0].text)                #pdf명
        pdf.append(j.find('a', class_='btnFile')['href'])   #pdf 링크
    pdf_url.append(pdf)
