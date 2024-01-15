import datetime
import random

import openpyxl
import requests
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pprint
from bs4 import BeautifulSoup
import json
import re
import smtplib  # SMTP 사용을 위한 모듈
from email.mime.multipart import MIMEMultipart  # 메일의 Data 영역의 메시지를 만드는 모듈
from email.mime.text import MIMEText  # 메일의 본문 내용을 만드는 모듈
from email.mime.base import MIMEBase
from email import encoders

def GetGoogleSpreadSheet():
    scope = 'https://spreadsheets.google.com/feeds'
    json = 'credential.json'
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json, scope)
    gc = gspread.authorize(credentials)
    sheet_url = 'https://docs.google.com/spreadsheets/d/1fYF8kQnhM8DCDpF7J6LjgT_c9suA-kSgi3vJOXvj6nQ/edit?pli=1#gid=1518561689'
    doc = gc.open_by_url(sheet_url)
    worksheet = doc.worksheet('미카엘리')

    #=================전체정보가져오기
    all_data=worksheet.get_all_records()
    #==================맨 밑행에 데이타 넣기
    # pprint.pprint(all_data)
    dataList=[]
    for data in all_data:
        productNo=data['네이버상품코드']
        productName=data['상품명']
        url = data['상품 링크']
        data={'productNo':productNo,'productName':productName,'url':url}
        dataList.append(data)
    pprint.pprint(dataList)
    return dataList

def GetInfo(url):
    cookies = {
        'WMONID': 'wCYqcx43ces',
        '_fbp': 'fb.1.1704289034744.1442023299',
        '_ga': 'GA1.1.192084949.1704289035',
        'imitation_data': '31%2C121%2C068',
        'LastestProduct': '736396%7C',
        'XSRF-TOKEN': 'eyJpdiI6IkFycEpzdW0yNVNnVEluOEE5bTR0NVE9PSIsInZhbHVlIjoid1VkM1Q5TjRzVFptRmE5UzNQODJidXdlWG9WS24rTXpKdXMwVmtJdmFKc1lJbE9DTWVvbnpNZVl6MldoUzdhOSIsIm1hYyI6IjE0ZGY1MGQ0NDlhZWIxMjljNjRjZmY1MjczN2M0ZGE2OThkMGMwMjRjMDRhYjFhNzEzMjE1YzBmZmFiMzlkZjMifQ%3D%3D',
        'nextokmallweb_session': 'eyJpdiI6ImdHMk9cL092cGxVc3VBcHF4UlJrT0NnPT0iLCJ2YWx1ZSI6Ik9tMkFvRlhhYWVYQWIxNjFPZGFoQUNaRm1tQkpoQkU0NjE0cGl6Nk5qdkk3YzZHWnZsM0IzdXlKRFJKV28zT2QiLCJtYWMiOiIwZTkzZDk2ZTcxODQ2NzMzOGIyYTMwNzhhZWMxNDY5ZGFhM2I3MDYwNWNmZTgyODZkNTRlMzJjMWY0Nzg2ZDZmIn0%3D',
        'SESSION_GUEST_ID': 'eyJpdiI6IkNJazE1a2xwV2NmbVk1U051Mlh5UXc9PSIsInZhbHVlIjoiWGdHTjFwMWVVYzhsN0V4ZTVOQUg5VEF3dGNTTjIwUTlpenZhNU9XeHhjZHgrbXdURFR3OUdZS2lVWmY2Q095MCIsIm1hYyI6ImUxOWRiNTcyNDZkYzlhZmM5OTMzYWI3YzAzYjlkMGI4NTM2YjQ2MzcyOWRhOGM4Yjg5N2UxMGViOTEwYzgxMDEifQ%3D%3D',
        '_ga_CW9NG23BGD': 'GS1.1.1704896555.4.0.1704896555.60.0.0',
        '_ga_Y9HS705BSQ': 'GS1.1.1704896555.4.0.1704896555.0.0.0',
        '_ga_4D8KD9470S': 'GS1.1.1704896555.4.0.1704896555.0.0.0',
    }

    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        # 'Cookie': 'WMONID=wCYqcx43ces; _fbp=fb.1.1704289034744.1442023299; _ga=GA1.1.192084949.1704289035; imitation_data=31%2C121%2C068; LastestProduct=736396%7C; XSRF-TOKEN=eyJpdiI6IkFycEpzdW0yNVNnVEluOEE5bTR0NVE9PSIsInZhbHVlIjoid1VkM1Q5TjRzVFptRmE5UzNQODJidXdlWG9WS24rTXpKdXMwVmtJdmFKc1lJbE9DTWVvbnpNZVl6MldoUzdhOSIsIm1hYyI6IjE0ZGY1MGQ0NDlhZWIxMjljNjRjZmY1MjczN2M0ZGE2OThkMGMwMjRjMDRhYjFhNzEzMjE1YzBmZmFiMzlkZjMifQ%3D%3D; nextokmallweb_session=eyJpdiI6ImdHMk9cL092cGxVc3VBcHF4UlJrT0NnPT0iLCJ2YWx1ZSI6Ik9tMkFvRlhhYWVYQWIxNjFPZGFoQUNaRm1tQkpoQkU0NjE0cGl6Nk5qdkk3YzZHWnZsM0IzdXlKRFJKV28zT2QiLCJtYWMiOiIwZTkzZDk2ZTcxODQ2NzMzOGIyYTMwNzhhZWMxNDY5ZGFhM2I3MDYwNWNmZTgyODZkNTRlMzJjMWY0Nzg2ZDZmIn0%3D; SESSION_GUEST_ID=eyJpdiI6IkNJazE1a2xwV2NmbVk1U051Mlh5UXc9PSIsInZhbHVlIjoiWGdHTjFwMWVVYzhsN0V4ZTVOQUg5VEF3dGNTTjIwUTlpenZhNU9XeHhjZHgrbXdURFR3OUdZS2lVWmY2Q095MCIsIm1hYyI6ImUxOWRiNTcyNDZkYzlhZmM5OTMzYWI3YzAzYjlkMGI4NTM2YjQ2MzcyOWRhOGM4Yjg5N2UxMGViOTEwYzgxMDEifQ%3D%3D; _ga_CW9NG23BGD=GS1.1.1704896555.4.0.1704896555.60.0.0; _ga_Y9HS705BSQ=GS1.1.1704896555.4.0.1704896555.0.0.0; _ga_4D8KD9470S=GS1.1.1704896555.4.0.1704896555.0.0.0',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'sec-ch-ua': '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    response = requests.get(url, cookies=cookies, headers=headers)
    soup=BeautifulSoup(response.text,'lxml')
    
    print("url:",url,"/ url_TYPE:",type(url))
    images=soup.find_all("img")
    isSoldOut="재고있음"
    for image in images:
        if image['src'].find("bx_soldout_rb2.jpg")>=0:
            isSoldOut="품절"
            break
    print("isSoldOut:",isSoldOut,"/ isSoldOut_TYPE:",type(isSoldOut))
    options=soup.find_all("tr",attrs={'name':'selectOption'})

    try:
        price=soup.find("div",attrs={'class':'last_price'}).find('span',attrs={'class':'price'}).get_text().replace(",","")
        # 쉼표를 제외한 연속된 숫자 찾기
        numbers = re.findall(r'\d+', price)[0]
        price=numbers
    except:
        price=""
    print("price:",price)
    
    optionPriceList=[]
    sizeList=[]
    colorList=[]
    for option in options:
        color=option.find_all('td')[0].get_text()
        colorList.append(color)
        # print("color:",color,"/ color_TYPE:",type(color))
        size = option.find_all('td')[1].get_text()
        # print("size:",size,"/ size_TYPE:",type(size))
        sizeList.append(size)
        optionPrice = option.find_all('td')[3].get_text().replace(",","")
        print("optionPrice:",optionPrice,"/ optionPrice_TYPE:",type(optionPrice))
        regex=re.compile("\d+")
        numbers=regex.findall(optionPrice)[-1]
        optionPrice=int(numbers)-int(price)
        optionPriceList.append(optionPrice)
        # print("optionPrice:",optionPrice,"/ optionPrice_TYPE:",type(optionPrice))
        # print("----------------------------------")
    if len(options)>=1:
        originPrice=int(price)-min(optionPriceList)
    else:
        originPrice = price
    print("originPrice:",originPrice,"/ originPrice_TYPE:",type(originPrice))

    optionListPrices="\n".join(str(num) for num in optionPriceList)
    print("optionListPrices:",optionListPrices,"/ optionListPrices_TYPE:",type(optionListPrices))

    optionListSizes = "\n".join(str(num) for num in sizeList)
    print("optionListSizes:",optionListSizes,"/ optionListSizes_TYPE:",type(optionListSizes))
    
    optionListColors = "\n".join(str(num) for num in colorList)
    print("optionListColors:",optionListColors,"/ optionListColors_TYPE:",type(optionListColors))

    result=[isSoldOut,optionListColors,optionListSizes,originPrice,optionListPrices]
    return result
    
def SendMail(filepath):

    smtp_server = 'smtp.naver.com'
    smtp_port = 587

    # 네이버 이메일 계정 정보
    username = 'mike102jiro@naver.com'  # 클라이언트 정보 입력
    password = 'Jan240109$$$'  # 클라이언트 정보 입력

    # receiver='wsgt17@naver.com'
    receiver='mike102jiro@naver.com'
    # receiver=email

    # username = 'hellfir2@naver.com'  # 클라이언트 정보 입력
    # password = 'dlwndwo1!'  # 클라이언트 정보 입력
    # =================커스터마이징
    try:
        to_mail = receiver
    except:
        print("메일주소없음")
        return

    # =================

    # 메일 수신자 정보
    to_email = receiver

    # 참조자 정보
    cc_email = 'ljj3347@naver.com'

    # 메일 본문 및 제목 설정
    contentList=[]

    content="\n".join(contentList)


    # MIMEMultipart 객체 생성
    timeNow=datetime.datetime.now().strftime("%Y년%m월%d일 %H시%M분%S초")
    msg = MIMEMultipart('alternative')
    msg["Subject"] = "[결과]OKMALL 상품 크롤링 (현재시각:{})".format(timeNow)  # 메일 제목
    msg['From'] = username
    msg['To'] = to_email
    msg['Cc'] = cc_email  # 참조 이메일 주소 추가
    msg.attach(MIMEText(content, 'plain'))

    # 파일 첨부
    part = MIMEBase('application', 'octet-stream')
    with open(filepath, 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={filepath}')
    msg.attach(part)

    # SMTP 서버 연결 및 로그인
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(username, password)
    # 이메일 전송 (수신자와 참조자 모두에게 전송)
    to_and_cc_emails = [to_email] + [cc_email]
    server.sendmail(username, to_and_cc_emails, msg.as_string())
    # SMTP 서버 연결 종료
    server.quit()
    print("전송완료")

        





# # ==========리스트가져오기
# inputList=GetGoogleSpreadSheet()
# with open('inputList.json', 'w',encoding='utf-8-sig') as f:
#     json.dump(inputList, f, indent=2,ensure_ascii=False)
#
# # ===========디테일가져오기
# with open ('inputList.json', "r",encoding='utf-8-sig') as f:
#     inputList = json.load(f)
#
#
# wb=openpyxl.Workbook()
# ws=wb.active
# columnName=['네이버상품번호','상품명','링크','재고여부(재고있음/품절)','구매가능색상','구매가능사이즈','기본판매가격','옵션별 가격']
# ws.append(columnName)
# timeNow=datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
# for index,inputElem in enumerate(inputList):
#     text="{}/{}번째 확인중...".format(index+1,len(inputList))
#     if len(inputElem['url'])==0:
#         print("없어서스킵")
#         continue
#     print(text)
#     infos=GetInfo(inputElem['url'])
#
#
#     data=[inputElem['productNo'],inputElem['productName'],inputElem['url']]
#     data.extend(infos)
#
#     print("data:",data,"/ data_TYPE:",type(data))
#     ws.append(data)
#     print("====================")
#     filepath='result_{}.xlsx'.format(timeNow)
#     wb.save(filepath)
#     time.sleep(random.randint(5,10)*0.1)

# SendMail(filepath)
SendMail('result_20240114_221601.xlsx')