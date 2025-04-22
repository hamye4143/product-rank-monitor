import requests
import json
import firebase_admin
from firebase_admin import credentials
from firebase_admin import db
import os
import sys
from datetime import datetime
import pandas as pd
from datetime import datetime

load_dotenv()


def downloadExcel(column1,column2,column3):
    try:

        # 엑셀 파일 만들기
        raw_data = {'상품명' : column1,
                '키워드' : column2,
                datetime.today().strftime("%Y-%m-%d"): column3} #리스트 자료형으로 생성

        dataframe = pd.DataFrame(raw_data).set_index(['상품명', '키워드'])
        dataframe.to_excel('상품순위_'+datetime.today().strftime("%Y%m%d%H%M%S")+'.xlsx', sheet_name=datetime.today().strftime("%Y-%m-%d"))

    except Exception as err:
        print('[ERROR] ',err)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


#db 연결 및 데이터 가져오기
def firebase_get_db():
    cred_path = os.getenv("FIREBASE_CREDENTIAL_PATH")
    db_url = os.getenv("FIREBASE_DB_URL")

    cred = credentials.Certificate(resource_path(cred_path))
    firebase_admin.initialize_app(cred, {'databaseURL': db_url})

    # 최초 데이터 업로드
    dir = db.reference()

    # dir.update({'상품':[
    #                     {
    #                     'code': '9057042500',
    #                     'shownName' : '오뚜기컵밥',
    #                     'keyword':['컵밥','오뚜기컵밥'],
    #                     },
    #                     {
    #                     'code': '9160871913',
    #                     'shownName' : 'CJ컵밥(미역국)',
    #                     'keyword':['컵밥','햇반컵반'],
    #                     },
    #                     {
    #                     'code': '9069617230',
    #                     'shownName' : '장어',
    #                     'keyword':['장어','민물장어'],
    #                     },
    #                     {
    #                     'code': '9143864716',
    #                     'shownName' : '꽃게',
    #                     'keyword':['꽃게','손질꽃게','냉동꽃게','절단꽃게'],
    #                     },
    #                     {
    #                     'code': '9095857460',
    #                     'shownName' : '고구마',
    #                     'keyword':['고구마','해남고구마','한입고구마','꿀고구마'],
    #                     },
    #                     {
    #                     'code': '9161193503',
    #                     'shownName' : '송편',
    #                     'keyword':['송편','모시송편'],
    #                     },
    #                     {
    #                     'code': '9352725346',
    #                     'shownName' : '에스더 6개월분',
    #                     'keyword':['글루타치온','여에스더글루타치온5x'],
    #                     },
    #                     {
    #                     'code': '9030816726',
    #                     'shownName' : '에스더 1개월분',
    #                     'keyword':['글루타치온','여에스더글루타치온5x'],
    #                     },
    #                     {
    #                     'code': '9058038522',
    #                     'shownName' : '비비고',
    #                     'keyword':['비비고 국','비비고 탕','비비고 육개장'],
    #                     },
    #                     {
    #                     'code': '9284001916',
    #                     'shownName' : '오뚜기 3분요리',
    #                     'keyword':['3분요리','3분카레','3분카레 순한맛'],
    #                     },
    #                     {
    #                     'code': '9199465503',
    #                     'shownName' : '스팸 선물세트',
    #                     'keyword':['스팸 선물세트'],
    #                     },
    #                     {
    #                     'code': '9261706647',
    #                     'shownName' : '에스티로더',
    #                     'keyword':['에스티로더 갈색병','에스티로더 갈색병 100ml'],
    #                     },
    #                     {
    #                     'code': '9275038471',
    #                     'shownName' : '피데기 오징어',
    #                     'keyword':['반건조오징어','피데기오징어'],
    #                     },
    #                     {
    #                     'code': '9215900158',
    #                     'shownName' : '말랑촉촉 수제간식',
    #                     'keyword':['훈련용간식','강아지노즈워크','강아지간식'],
    #                     },
    #                     {
    #                     'code': '9336703264',
    #                     'shownName' : '비비고죽 소고기죽CJ 죽',
    #                     'keyword':['비비고죽','즉석죽'],
    #                     },
    #                     {
    #                     'code': '9275038471',
    #                     'shownName' : '반건조 오징어',
    #                     'keyword':['반건조오징어','피데기오징어'],
    #                     },
    #                     {
    #                     'code': '9570652916',
    #                     'shownName' : '과메기',
    #                     'keyword':['과메기','구룡포과메기'],
    #                     },
    #                     {
    #                     'code': '9419042824',
    #                     'shownName' : '가리비',
    #                     'keyword':['가리비','통영가리비'],
    #                     },
    #                     {
    #                     'code': '9453852439',
    #                     'shownName' : '굴',
    #                     'keyword':['통영굴','석화'],
    #                     },
    #                     {
    #                     'code': '9429452754',
    #                     'shownName' : '쭈꾸미',
    #                     'keyword':['쭈꾸미볶음','쭈꾸미양념'],
    #                     },
    #                     {
    #                     'code': '9531511991',
    #                     'shownName' : '귤',
    #                     'keyword':['귤','감귤'],
    #                     },
    #                     {
    #                     'code': '9601617528',
    #                     'shownName' : '홍게',
    #                     'keyword':['홍게','대게'],
    #                     },
    #                   ]
    #             }) #데이터 업데이트

    datas = dir.get()
    return datas


#랭킹을 가져온다
def getRanks(keyword, mallId):
    seller='포레스트뉴'
    pagingIndex=1
    is_top=False
    product_name = ''
    returnValue=''

    headers = {
        'authority': 'search.shopping.naver.com',
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        # 'cookie': 'NNB=D3SUWNPSZBMWI; SHP_BUCKET_ID=7; NSCS=1; ASID=77c561820000018829e0cf670000004b; _ga_GDZF0H5BMH=GS1.1.1685580047.1.1.1685581183.60.0.0; autocomplete=use; NV_WETR_LOCATION_RGN_M="MDkxNDAxMDQ="; NV_WETR_LAST_ACCESS_RGN_M="MDkxNDAxMDQ="; nx_ssl=2; _ga_6Z6DP60WFK=GS1.2.1688975956.1.0.1688975956.60.0.0; nid_inf=-1407962044; NID_AUT=f8b+pILdyEifcBSTncJoLFl0+Whc/nOq54RnK6lJaAaDSQrTM8rP3X9oF+VA8xzm; NID_JKL=VmXShvyJ0iTA0i695jOR318HEyMvP9VR8/oyWhXy/bA=; _ga=GA1.2.1701136973.1690257791; ncpa=4536612|lkmruuxc|163b9bb0a6fc9e23efb55cdabd2a0d5b9a10fd9a|s_2da8726e5e7f6|8cb9022f652af4543fe1aa6064e8f833cfbd55e1; _ga_ZGQY5GH55D=GS1.1.1690562927.14.0.1690562927.0.0.0; NID_SES=AAABgnulEvEtdAPHwJ1sg5+M/yRdk1ri/MgNfPd4e9PNops4pWmCXi2zGt5QparrNe6gaGk0Cl/x3lR69AtOVPSsimjkZhuKpQNaQrkwMRUqBvRyom7MQ0DbubjA+mWSQ81NWinOJcvdLHxpR/VzaCZ9pJKX+hop9pFca0u8/VbGZfBSD/t7nIZ1GJxO5MD+gKvlF4MGLLh5PE5KmMORZJG3RJxqFnvsOKGQV4ekMjYP6cOLw4KBPpBllm6G+FCCRl8neP8XPK133d/K1tqd7rKNZeNHqthMVvZfE9vf97y6U7VDOMfsNjRRe8/XPyKuxgLiRYIeF8Z8spoiKOU6WtgVgrOJaPFnJ/u9yaPH9tbQxLrWBmCwRvNLJ7cb5s0jkRoA3E64bGYC+d/MlFtO04hT4xaJhhMbyeVAK9NwS/Z+mLXUnLSEMfdi+PfxB+SaRs6joH6TZfjbmrf67kWDwRpr60AVemusvpk/HdfyaSd6k0rTp4yf9GEP40H919frAVq4eblKLYKBLPMIKqy16LUt/Xk=; spage_uid=',
        'logic': 'PART',
        'referer': 'https://search.shopping.naver.com/search/all?query=%EC%98%A4%EB%8B%88%EC%B8%A0%EC%B9%B4%ED%83%80%EC%9D%B4%EA%B1%B0&cat_id=&frm=NVSHATC',
        'sbth': 'b77ad4005ee821024462d4b57e59d24e95ccc7062f65af77e4baf2a0c70a1c6677c58d9032794b7a0f06ee8fd74bcdde',
        'sec-ch-ua': '"Whale";v="3", "Not-A.Brand";v="8", "Chromium";v="114"',
        'sec-ch-ua-arch': '"x86"',
        'sec-ch-ua-bitness': '"64"',
        'sec-ch-ua-full-version-list': '"Whale";v="3.21.192.18", "Not-A.Brand";v="8.0.0.0", "Chromium";v="114.0.5735.138"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-model': '""',
        'sec-ch-ua-platform': '"Windows"',
        'sec-ch-ua-platform-version': '"10.0.0"',
        'sec-ch-ua-wow64': '?0',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Whale/3.21.192.18 Safari/537.36',
    }

    chk=False

    try:
        while pagingIndex <=5:
            params = {
            'eq': '',
            'frm': 'NVSHATC',
            'iq': '',
            'origQuery': keyword,
            'pagingIndex': str(pagingIndex),
            'pagingSize': '40',
            'productSet': 'total',
            'query': keyword,
            'sort': 'rel',
            'viewType': 'list',
            'xq': '',
            }
            response = requests.get('https://search.shopping.naver.com/api/search/all', params=params, headers=headers)
            result = json.loads(response.text)

            product_data = result['shoppingResult']['products']

            for idx, product in enumerate(product_data):
                #price = product['price']
                mallName= product['mallName']

                if mallName == seller and product['mallProductId'] == mallId :#셀러라면
                    chk=True
                    product_name=product['productTitle']
                    if idx+1 ==1: # 맨 상단에 있다면
                        is_top = True
                    break #for문 나가기

            if chk==True:
                break

            pagingIndex+=1


        if chk == False:
            returnValue = '5p 이후'
            print('[',keyword,']',': 5p 이후')
        else:
            if is_top == True:
                print('[', keyword,']-','(',product_name,'):',pagingIndex,'p 맨 상단')
                returnValue = str(pagingIndex)+'p 맨 상단'
            else:
                print('[', keyword,']-','(',product_name,'):',pagingIndex,'p')
                returnValue = str(pagingIndex)+'p'

        return returnValue

    except Exception as err:
        print('[', keyword, '] 에러가 발생했습니다.')
        return err


def main():

    #db에서 데이터 가져오기
    datas= firebase_get_db()
    # excel_result=[]
    pagingIndex = ''
    column1=[]
    column2=[]
    column3=[]

    for data in datas['상품']:
        for keyword in data['keyword']:
            pagingIndex = getRanks(keyword, data['code'])
            # excel_result.append([data['shownName'], keyword, pagingIndex])
            column1.append(data['shownName'])
            column2.append(keyword)
            column3.append(pagingIndex)

    # print(excel_result)

    print('--------------------끝--------------------')

    # downloadExcel(datas['상품'])
    downloadExcel(column1,column2,column3)


    # print('resource_path',resource_path('firebase_key.json'))
    #pyinstaller --icon=icon.ico --onefile --add-data "firebase_key.json;." ranking_system.py
    os.system("pause")

if __name__ == "__main__":
    main()
