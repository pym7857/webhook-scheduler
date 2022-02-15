#python 3.8.6
import requests
import json
import pandas as pd 
import datetime
from korean_lunar_calendar import KoreanLunarCalendar

if __name__ == "__main__":

    # 컬럼, 행 쳐내기 
    df = pd.read_excel('data/birthday.xlsx', names=['dd', 'dept', 'name', 'birthday'])
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
        year_ = 2022
        month_ = int(date.split('월')[0])
        day_ = int(date.split('월')[1].split('일')[0])

        calendar = KoreanLunarCalendar()
        calendar.setLunarDate(year_, month_, day_, False)
        #print(calendar.SolarIsoFormat())
        return calendar.SolarIsoFormat()

    # 양력이면 정해진 형식으로, 음력이면 양력형식으로
    def isLunar(x):
        if str(x).split("(")[1].split(")")[0] == "음":
            new_date = convert_to_korean_date(str(x).split("(")[0])
            #print(new_date)
            return new_date # 음력 
        return '2022-' + str(x).split("월")[0] + '-' + str(x).split("월")[1].strip().split("일")[0] # 양력 

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
    # now = datetime.datetime.now()
    # nowDate = now.strftime('%Y-%m-%d')
    # print(nowDate)      # 2018-07-28

    nowDate = '2022-4-4'

    # TODO
    url = 'https://lselectricdwp.webhook.office.com/webhookb2/67573c49-c614-4240-908a-bc778b3855a9@247258cc-5eb2-4fd4-9bb2-f272103f0c34/IncomingWebhook/a2f13b2f58de4f7eac433fb7f5a73e63/667762b1-4f18-4ffe-bf9c-3b0df246edc0'

    # 현재날짜와 일치하는 사람의 부서, 이름 추출 
    if nowDate in dx_dict.keys():
        for dept, name in dx_dict[nowDate]:
            payload = {
                "text": f"{dept} {name} 생일이에요! 축하해요!"
            }
            headers = {
                'Content-Type': 'application/json'
            }
            response = requests.post(url, headers=headers, data=json.dumps(payload))
            print(response.text.encode('utf8'))


    # url = 'https://lselectricdwp.webhook.office.com/webhookb2/67573c49-c614-4240-908a-bc778b3855a9@247258cc-5eb2-4fd4-9bb2-f272103f0c34/IncomingWebhook/a2f13b2f58de4f7eac433fb7f5a73e63/667762b1-4f18-4ffe-bf9c-3b0df246edc0'
    # payload = {
    #     "text": "Sample alert text"
    # }
    # headers = {
    #     'Content-Type': 'application/json'
    # }
    # response = requests.post(url, headers=headers, data=json.dumps(payload))
    # print(response.text.encode('utf8'))
