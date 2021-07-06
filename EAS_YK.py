# -*- coding:utf-8 -*-
from tkinter import *;
import pandas as pd
import tkinter.messagebox as tm
import tkinter.filedialog
import datetime
import traceback
#import xlrd
import os
import cx_Oracle as oracle
import requests
import json
from suds.client import Client



tiw=Tk("EAS英克组合小程序0.1");
tiw.title();
tiw.geometry("150x50");
menubar=Menu(tiw)
content=[['煎饼提示:选择功能后请按照提示选择源文件！']]
Main=['煎饼组合0714']
for i in range(len(Main)):
    #新建一个空的菜单,将menubar的menu属性指定为filemenu，即filemenu为menubar的下拉菜单
    filemenu = Menu(menubar, tearoff=0)
    for k in content[i]:
        filemenu.add_command(label = k)
    menubar.add_cascade(label=Main[i], menu=filemenu)

# 将root的menu属性设置为M
tiw['menu'] = menubar
#tiw.mainloop()

a = LabelFrame(tiw, height=22, width=50, text='组合功能')
a1 = LabelFrame(tiw, height=22, width=50, text='组合功能')
a.pack(side='left', anchor='ne')
a1.pack(side='left', anchor='ne')

def appendStr99():  #webservice 链接 英克 WMS
 try:
     # github_url = 'http://60.12.218.220:6694/ormrpc/services/EASLogin?wsdl'

     github_url = 'http://60.12.218.220:60011/wmswebservice/services/pubservice?wsdl'
     client=Client(github_url)
     print(client)
     data = json.dumps({"companystype": "1",
                        "erpcompanyid": "2",
                        "goodsownerid": "21",
                        "orgno": "1",
                        "Warehid":"4"
                        })
     print(client.service.pubIntf(type=1001,jsonstr=data))





     # data = json.dumps({"companystype": "1",
     #                    "erpcompanyid": "2",
     #                    "goodsownerid": "21",
     #                    "orgno": "1",
     #                    "Warehid":"4"
     #
     #                    })

     # r = requests.post(github_url, data, auth=('user', 'kduser'))
     # r = requests.post(github_url,data)
     # print(r.status_code)
     # print(r.text)

 except Exception as error:
      tm.showerror(title="煎饼提示前方路堵",
              message="请检查提交源文件是否正确 '" + str(error) + "'.",
               detail=traceback.format_exc())
def appendStr98():  #测试组
 try:
    os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'

    os.environ['path'] = r'C:\instantclient_19_3'

    # 查询操作（查）

    db = oracle.connect('heseas/kingdee@60.12.218.220:6694/ORCLEAS')  # 数据库连接
    cursor = db.cursor()  # 创建cursor
    cursor.execute("SELECT * FROM t_gl_acctcussent")  # 执行sql语句
    rs = cursor.fetchall()  # 一次返回所有结果集 fetchall
    id2 = rs[0][0]  # 去除多余的内容
    print(id2)  # 打印内容

    db.close()  # 关闭数据库连接

    tkinter.messagebox.showinfo(id2, "下载地址：待定");


 except Exception as error:
      tm.showerror(title="煎饼提示前方路堵",
              message="请检查提交源文件是否正确 '" + str(error) + "'.",
               detail=traceback.format_exc())
appendBtn=Button(a,text="EAS",width=22,height=1,command=appendStr98);
appendBtn=Button(a1,text="WMS",width=22,height=1,command=appendStr99);
appendBtn.pack();

tiw,mainloop();