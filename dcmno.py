import OpenDartReader
import requests                     #post로 dcmno 요청하기 위함
from bs4 import BeautifulSoup       #html로 쉽게 파싱하기 위함.

#홍차장님 증정
api_key='8fd02dac927493a7161e13a34e78062b49197e59'
dart = OpenDartReader(api_key)

#회사 코드 먼저 가져오자. 3개월씩 제한이 있음
#dart.list()이 반환한 pandas dataframe을 list로 변환하여 사용.
comp_df=dart.list(start='2023-12-13',end='2024-02-13',kind='G')
li=comp_df.values.tolist()
li.insert(0,comp_df.columns.values.tolist())

#자, 이 짓을 10000번 반복하자. (10000번은 혹시모르니 냄겨둠)

dart.find_corp_code('00267526')


'''
dcmno는 reportForm으로 post요청을 하여 받아온다.
서버사이드 렌더링이므로 클라이언트 측에서 소스를 직접 볼 수 없다.
따라서 requests 모듈로 http를 요청하여 받아온 html 내 dcmno를 파싱한다.

참고) 폼의 구성은 다음과 같다.
<form name="reportForm" method="post" action="/dsaf001/main.do">
	<input type="hidden" name="rcpNo">
	<input type="hidden" name="dcmNo">
	<input type="hidden" name="keyword">
	<input type="submit" style="display:none;">
</form>

requests 모듈에서 post로 날리는 법을 알게되면
아래처럼 번거롭게파싱할 필요없이 dcpno를 받아올 수 있다.
우선은 get으로 html 소스 전체를 받아와서 파싱하도록하자.
'''

#post 날리기
datas={
    'dcmNm':'',
    'tagId':''
}

url='/dsaf001/main.do?rcpNo=20240206000062'
post_response = requests.post(url,data=datas)

#파싱을 시작한다.
#rcpno를 연결하여 url을 생성한다.
url='https://dart.fss.or.kr/dsaf001/main.do?rcpNo=20240201000567'
#해당 url을 get으로 js렌더링이 된 html을 받아온다.
response=requests.get(url)
#다운로드 버튼 내에 dcmno가 숨어있다. css selector으로 다운로드 파일의 위치를 잡는다.
path_btn='body > div.wrapper > div > div.header > div.bottom > div.rightWrap > button.btnDown'

#혹시 모를 거절 요청 대비
if response.status_code == 200:
    html = response.text
    soup = BeautifulSoup(html, 'html.parser')
else:
    print('뭔가 잘못되었습니다')

target=soup.select(path_btn)
#여기서 문자열을 파싱하여 dcmno를 얻는다.
target=str(target).split()[3]
target=target.split(';')[0].replace('\'','')
target=int(target[:-1]) #맨 뒤의 )를 지우고 숫자로 바꾼다.
