# -*- coding: utf-8 -*-
import pymysql
import cx_Oracle as oracle
import os
import requests
import json

data = {
    "application_id": 1000001,
    "application_secret": "8888",
    "expired": 31536000
}
headers = {'Content-Type': 'application/json',
           "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36", }
###写入表头并POST
response = requests.post(url='https://api.nb-health.com/v2/ticket', json=data, headers=headers);
##返回信息
print(response.text)
df = response.text
# 将JSON数据解码为dict（字典）
json_str = json.loads(df)
print(type(json_str))  ###查询数组格式  dict 字典
##取出ticket
print(json_str['ticket'])

###拼接表头
headers_hbs = {'Connection': 'keep-alive',
               'Pragma': 'no-cache',
               'Cache-Control': 'no-cache',
               'Accept': 'application/json',
               'Accept-Language': 'zh-CN',
               'X-Huoban-Ticket': json_str['ticket'],
               'X-Huoban-Return-Alias-Space-Id': '3300000000014532',
               "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36"
               }
# print(headers_hbs)

###新增地址
url_add = "https://api.nb-health.com/v2/item/table/2100000020012945/create"
# url_add="https://api.nb-health.com/v2/item/table/1888024/create"

os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
os.environ['path'] = r'c:\instantclient_19_3'

# 查询操作（查）

db = oracle.connect('heseas', '9999', '172.16.24.223:1521/eas99')
cursor = db.cursor()  # 创建cursor

sql = "SELECT a.fname_l2 客户名称,b.fname_l2 集团名称,ach.cfykcustomer 英克代码,a.fnumber 金蝶编码,c.FNUMBER 集团编码,a.FSIMPLENAME 类型 FROM CT_CUS_CustomerMapping ach " \
      "inner join t_bd_customer a on ach.cfeascustomerid=a.fid inner join t_bd_generalasstacttype b on ach.cfeasjtid=b.fid " \
      "inner join t_bd_generalasstacttype c on ach.cfeasjtid=c.fid where ach.cfykcustomer is not null order by ach.cfykcustomer"
cursor.execute(sql)


item_list = []
item_list1 = []

if cursor.fetchone:  # 返回值是单个的元组,也就是一行记录,如果没有结果,那就会返回null
    for row2 in cursor.fetchall():  # 返回值是多个元组,即返回多个行记录,如果没有结果,返回的是()
        # print(row2)
        new_kemc = str(row2[0])  # 0是名字
        # new_wtfid = (row2[1])
        new_ykid = (row2[2])  # 英克ID
        new_easbm = (row2[3])  # EAS编码
        # new_djbh = (row2[4])
        # new_djlx = (row2[5])
        # print(new_kemc,new_ykid,new_easbm)

        data1 = {"fields": {
            "2200000150418434": new_kemc,  ##字段写入
            "2200000150418435": new_easbm,
            "2200000150418436": new_ykid}}

        item_list.append(data1)
        item_list1.append(data1)
        print(len(item_list))

        if len(item_list) == 200:
            response1 = requests.post(url=url_add, json={"items": item_list}, headers=headers_hbs);
            print(response1)
            item_list.clear()
        else:
            continue

    # 200外的数据区分写入
    print(len(item_list1) / 200)
    item_list2 = []
    # item_list3=[]
    b = round(len(item_list1) / 200 - int(len(item_list1) / 200), 4)
    c = (b * -200)
    print(b, int(c))
    print({"items": item_list1[int(c):-1]})
    item_list2 = item_list1[int(c):-1]
    print(item_list1[-1])  # 补上最后一条
    item_list2.append(item_list1[-1])
    # item_list3 = item_list1[-1]
    response2 = requests.post(url=url_add, json={"items": item_list2}, headers=headers_hbs);
    print(response2)



