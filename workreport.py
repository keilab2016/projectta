import gspread
from oauth2client.service_account import ServiceAccountCredentials

import re


def getAllData(spreadsheet_key, sheet_name):
    #jsonファイルを使って認証情報を取得
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    c = ServiceAccountCredentials.from_json_keyfile_name('projectta2021-d0f50421fa21.json', scope)

    #認証情報を使ってスプレッドシートの操作権を取得
    gs = gspread.authorize(c)

    #共有したスプレッドシートのキー（後述）を使ってシートの情報を取得
    worksheet = gs.open_by_key(spreadsheet_key).worksheet(sheet_name)
    list_of_lists = worksheet.get_all_values()

    #print(list_of_lists)
    return list_of_lists

def findData(userid, month, list_of_lists):
    alldata={}
    for line in list_of_lists:
        if line[0] == '学籍番号':
            ids = line
        elif '氏名' in line[0] or '名前' in line[0]:
            names = line
        elif '研究室' in line[0]:
            labs = line
        else:
            result = re.findall(r'([0-9]+)/([0-9]+)', line[0])
            if len(result)==1 and int(result[0][0])==month:
                alldata[result[0][0]+'/'+result[0][1]]=line

    #print(ids)
    #print(names)
    print(alldata)
    data={}
    if userid in ids:
        idx = ids.index(userid)
        data['id']=ids[idx]
        data['name']=names[idx]
        data['lab']=labs[idx]
        for day,line in alldata.items():
            data[day]=line[idx]
    else:
        print(userid,'not found in',ids)
    return data



list_of_lists = getAllData('102Fbm3HKFOeCg6Q0CEAQE88CLQ3xvtW3bTPDGydx02A',"2021前期")
data=findData('1018078', 5, list_of_lists)
