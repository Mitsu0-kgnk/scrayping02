import requests
import os
import openpyxl
import pandas as pd 

REQUEST_URL = 'https://app.rakuten.co.jp/services/api/IchibaItem/Search/20170706'
REQUEST_URL_2 = 'https://app.rakuten.co.jp/services/api/IchibaItem/Ranking/20170628'
APP_ID = '1085838838124753303'

kw = input('')


df = pd.DataFrame()
l = ['itemName', 'itemPrice','reviewCount','reviewAverage','itemUrl', 'shopName','genreId']
d = {}

params = {
    'format':'json',
    'applicationId':APP_ID,
    'keyword':kw,
    'page':1,
    'sort':'-reviewCount'
}

res = requests.get(REQUEST_URL,params)
result = res.json()

items = result['Items']

for i in range(len(items)):


    item = items[i]['Item']


    
    pics = item['mediumImageUrls']

    num = len(item['mediumImageUrls'])

    pic_url = []

    for i in range(num):

        pic_url.append(pics[i]['imageUrl'])

    for i,url in enumerate(pic_url):
        res = requests.get(url)
        image = res.content
        file_name = item['itemName'][0:6]+'_{}.jpg'.format(i)
        file_name = ''.join(filter(str.isalnum, file_name)) 
        file_name = file_name+'.jpg'
        with open(r'C:\Users\NDY02\Desktop\Python\Scrayping\pics\{}'.format(file_name),'wb')as f:
            f.write(image)
            
    for k in l:
        d2 = item.get(k)
        d[k] = d2
        _df = pd.DataFrame(d,index=[i])
    df = df.append(_df)

df = df.sort_values(['reviewCount','reviewAverage'],ascending=[False,False])

with pd.ExcelWriter('楽天_{}.xlsx'.format(kw)) as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)

a = df['genreId'].mode()

params2 = {
    'format':'json',
    'applicationId':APP_ID,
    'genreId':a
}

res2 = requests.get(REQUEST_URL_2,params2)
result2 = res2.json()

l2 = ['rank','itemName','itemPrice','itemUrl','reviewCount','reviewAverage']
d3 = {}
df3 = pd.DataFrame()

for i in range(len(result2['Items'])):
    r = result2['Items'][i]['Item']
    for k in l2:
        v = r.get(k)
        d3[k] = v
        _df = pd.DataFrame(d3,index=[i])
    df3 = df3.append(_df)

wb = openpyxl.load_workbook('楽天_{}.xlsx'.format(kw))
ws = wb.create_sheet(title='Sheet2')

with pd.ExcelWriter('楽天_{}.xlsx'.format(kw), mode='a', engine="openpyxl") as writer:
    df3.to_excel(writer, sheet_name='Sheet2',index=False)
