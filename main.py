#python 3.9.1 (mac)
import requests
import json
import pandas as pd 
import datetime
from korean_lunar_calendar import KoreanLunarCalendar
import pymsteams

if __name__ == "__main__":

    now_year = '2022'

    # 컬럼, 행 쳐내기 !
    df = pd.read_excel('data/birthday_7.xlsx', names=['dd', 'dept', 'name', 'birthday'])
    df.drop(['dd'], axis=1, inplace=True)
    df = df[1:]

    # TODO: df 부서 nan 인것들 채우기 
    lst = df.loc[df['dept'].notna()].index.tolist()
    i = 0 
    rows = []
    for idx, row in df.iterrows():
        if idx in lst:
            i = idx 
            rows.append(row)
        if idx not in lst:
            row['dept'] = df.loc[i]['dept']
            #print(row)
            rows.append(row)
    #print(rows)
    res = pd.DataFrame(rows)
    #print(res)

    # 음력 -> 양력 
    def convert_to_korean_date(date): # 1월 2일 
        #print(date)
        month_ = int(date.split('월')[0])
        day_ = int(date.split('월')[1].split('일')[0])

        calendar = KoreanLunarCalendar()
        calendar.setLunarDate(int(now_year), month_, day_, False)
        #print(calendar.SolarIsoFormat())
        return calendar.SolarIsoFormat()

    # 2022-4-4  →  2022-04-04
    def plus_zero(date):
        c = date.split('-')[0]
        a = date.split('-')[1]
        b = date.split('-')[2]
        if len(a) == 1:
            a = '0' + a
        if len(b) == 1:
            b = '0' + b
        return c + '-' + a + '-' + b

    # 양력이면 정해진 형식(xxxx-xx-xx)으로, 음력이면 양력으로 바꾼다음 형식(xxxx-xx-xx)통일
    def isLunar(x):
        x = x.strip()
        if str(x).split("(")[1].split(")")[0] == "음": # 음력
            new_date_lunar = convert_to_korean_date(str(x).split("(")[0])
            return plus_zero(new_date_lunar)  
        new_date_solar = str(now_year) + '-' + str(x).split("월")[0] + '-' + str(x).split("월")[1].strip().split("일")[0] # 양력 
        return plus_zero(new_date_solar)

    res['solar'] = res['birthday'].apply(isLunar)
    print(res)

    # dict에 집어넣기 
    dx_dict = dict()
    for idx, row in res.iterrows():
        if row['solar'] in dx_dict.keys():
            dx_dict[row['solar']].append( (row['dept'], row['name']) )
        else:
            dx_dict[row['solar']] = [(row['dept'], row['name'])]
    print(dx_dict)

    # TODO: 실사용때 여기구문 사용할것 
    # 현재 날짜 추출 
    now = datetime.datetime.now()
    nowDate = str(now.strftime('%Y-%m-%d'))
    print(nowDate)      # 2018-07-28

    #nowDate = '2022-03-19'

    # TODO
    
    # 비전실 일반
    web_hook_url = 'https://lselectricdwp.webhook.office.com/webhookb2/6092c94e-c390-4d39-808d-67fff38a44e7@247258cc-5eb2-4fd4-9bb2-f272103f0c34/IncomingWebhook/22c6a1caa8ad4570befbb61fb61f5bec/667762b1-4f18-4ffe-bf9c-3b0df246edc0'
    
    # dx-lab 
    #web_hook_url = 'https://lselectricdwp.webhook.office.com/webhookb2/67573c49-c614-4240-908a-bc778b3855a9@247258cc-5eb2-4fd4-9bb2-f272103f0c34/IncomingWebhook/d2590810eae4454abf56af9d1579fe05/667762b1-4f18-4ffe-bf9c-3b0df246edc0'
    
    # playground 
    # web_hook_url = 'https://lselectricdwp.webhook.office.com/webhookb2/67573c49-c614-4240-908a-bc778b3855a9@247258cc-5eb2-4fd4-9bb2-f272103f0c34/IncomingWebhook/a2f13b2f58de4f7eac433fb7f5a73e63/667762b1-4f18-4ffe-bf9c-3b0df246edc0'

    myTeamsMessage = pymsteams.connectorcard(web_hook_url)

    # 현재날짜와 일치하는 사람의 부서, 이름 추출 
    if nowDate in dx_dict.keys():
        for dept, name in dx_dict[nowDate]:
            teams_message = pymsteams.connectorcard(web_hook_url)

            # create the section
            myMessageSection = pymsteams.cardsection()
            # Section Title
            myMessageSection.title("📧생일축하 메세지📧")
            # Activity Elements
            myMessageSection.activityTitle(f"{dept}")
            myMessageSection.activitySubtitle(f"{name}")
            myMessageSection.activityImage("https://upload.wikimedia.org/wikipedia/commons/thumb/8/8e/LS_logo.svg/1200px-LS_logo.svg.png")
            # myMessageSection.activityText("Congraturation!")
            # Facts are key value pairs displayed in a list.
            # myMessageSection.addFact("this", "is fine")
            # myMessageSection.addFact("this is", "also fine")
            # Section Text
            myMessageSection.text("HAPPY BIRTHDAY !")
            # Section Images
            myMessageSection.addImage("https://i.imgur.com/0AVmRiD.jpeg", ititle="This Is Fine")
            # Add your section to the connector card object before sending
            teams_message.addSection(myMessageSection)

            teams_message.text(f"🎉오늘은 {dept} {name}님의 생일이에요! 축하해요!")
            teams_message.color('#DD7AF5') # theme color 
            teams_message.send()
