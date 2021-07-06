# -*- coding:utf-8 -*-
import  tkinter  #导入Tkinter
from tkinter import *;
import tkinter.messagebox as tm
from smtplib import SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import make_header
import os
import traceback
import pandas as pd
import tkinter.filedialog
import datetime

import xlrd

# import  HP_tk  as  htk   #导入htk

#创建Tkinterz主窗口
tiw=Tk();

tiw.title("基因邮件助手0519");
tiw.geometry("250x95");
menubar=Menu(tiw)
content=[['1.分解对账单 2.分解未开票 3.点击发送邮件 版本V1.3']]
Main=['使用说明']
for i in range(len(Main)):
    #新建一个空的菜单,将menubar的menu属性指定为filemenu，即filemenu为menubar的下拉菜单
    filemenu = Menu(menubar, tearoff=0)
    for k in content[i]:
        filemenu.add_command(label = k)
    menubar.add_cascade(label=Main[i], menu=filemenu)

# 将root的menu属性设置为M
tiw['menu'] = menubar
a = LabelFrame(tiw, height=22, width=50)
a.pack(side='left', anchor='ne')

def appendStrJY01():  #基因测试组
 try:
    tkinter.messagebox.showinfo("提醒", "请先选择基因导出客户对账单");
    data1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径  str(datetime.datetime.now().strftime('%Y%m%d'))

    # data["发票日期1"]=datetime.strptime(data["发票日期"], "%Y%m%d")
    # data = pd.datetime.strftime(data1.iloc[I, 4], "%Y-%m-%d")
    # print(data)
    ##修改表内日期，去掉时间
    data1["发票日期1"] = pd.datetime.strftime(data1.iloc[0, 4], "%Y-%m-%d")
    data1["应到款日期1"] = pd.datetime.strftime(data1.iloc[0, 8], "%Y-%m-%d")

    data2 = data1.drop(["发票日期","应到款日期"], axis=1)  # 删列
    ####新增小计

    data3 = data2.groupby(["客户名称"], as_index=False)["欠款（逾期应收账款）","应收账款（未逾期，提醒待追回）"].sum();
    data3["大区"] = ""
    data3["销售"] = ""
    data3["客户编号"] = ""
    data3["发票日期1"] = ""
    data3["发票号码"] = ""
    data3["求和项:原币应收金额"] = ""
    data3["账期"] = ""
    data3["应到款日期1"] = ""

    data5 = pd.concat([data2, data3], ignore_index=True)
    data = data5[["大区","销售","客户编号","客户名称","发票日期1","发票号码","求和项:原币应收金额","账期","应到款日期1","欠款（逾期应收账款）","应收账款（未逾期，提醒待追回）"]]

    rows = data.shape[0]  # 获取行数 shape[1]获取列数
    department_list = []
    for i in range(rows):
       temp = data["客户名称"][i]
       if temp not in department_list:
          department_list.append(temp)  # 将客户名称存在一个列表中
    for department in department_list:
       new_df = pd.DataFrame()
       for i in range(0, rows):
          if data["客户名称"][i] == department:

             new_df = pd.concat([new_df, data.iloc[[i], :]], axis=0, ignore_index=True)
       # new_df.to_excel(str(department) + "对账单.xls", sheet_name=department, index=False)  # 将每个销售部门存成一个新excel
       new_df.to_excel(str(department) + "对账单.xls",index=False)  # 将每个销售部门存成一个新exceli
    tkinter.messagebox.showinfo("运行结果", "客户对账单整理成功！");


 except Exception as error:
      tm.showerror(title="煎饼提示前方路堵",
              message="请检查提交源文件是否正确 '" + str(error) + "'.",
               detail=traceback.format_exc())

def appendStrJY02():  #基因测试组  210413反馈加不了客户，采取用客户档案文件做拼接

 try:
    tkinter.messagebox.showinfo("提醒", "请先选择基因导出未开票");
    df1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径

    data1 = df1[df1["开票"] == "N"];

    tkinter.messagebox.showinfo("提醒", "请先选择基因客户档案表");
    df2 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径

    # df3 = df2.drop(df2.columns[[[[[[[[[[[0, 3, 4, 5, 7, 8, 10, 11, 12, 13, 17]]]]]]]]]]], axis=1);
    df3 = df2.groupby(["客户编号","客户全称"],as_index=False)["分店数"].sum();
    df4 = df3.rename(columns={'客户全称': '客户名称'});
    df5 = df4.drop(["分店数"],axis = 1)#删列

    df5["发货日期1"] = pd.datetime.strftime(data1.iloc[0, 7], "%Y-%m-%d")

    data10 = pd.merge(data1, df5, how='left', on=['客户编号']);
    #####新增小计

    data2 = data10.groupby(["客户名称"], as_index=False)["应收-未开票","销货数量", "本币税额", "本币税前金额","本币价税合计"].sum();
    df11 = data10.drop(["客户简称"], axis=1)  # 删列

    df12 = pd.concat([df11, data2], ignore_index=True)

    
    data = df12[
        ["业务员", "发票种类", "品  号", "品  名", "地区", "客户单号", "客户名称", "客户编号", "开票", "发货日期1","本币价税合计", "本币税前金额",
         "本币税额","销货数量","未开票数量","应收-未开票"
         ]]
    rows = data.shape[0]  # 获取行数 shape[1]获取列数
    department_list = []
    for i in range(rows):
        temp = data["客户名称"][i]
        if temp not in department_list:
            department_list.append(temp)  # 将客户名称存在一个列表中
    for department in department_list:
        new_df = pd.DataFrame()
        for i in range(0, rows):
            if data["客户名称"][i] == department:
                new_df = pd.concat([new_df, data.iloc[[i], :]], axis=0, ignore_index=True)
        # new_df.to_excel(str(department) + "对账单.xls", sheet_name=department, index=False)  # 将每个销售部门存成一个新excel
        new_df.to_excel(str(department) + "未开票.xls", index=False)  # 将每个销售部门存成一个新excel

    tkinter.messagebox.showinfo("运行结果", "客户未开票整理成功！");
    # df3.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630



 except Exception as error:
      tm.showerror(title="煎饼提示前方路堵",
              message="请检查提交源文件是否正确 '" + str(error) + "'.",
               detail=traceback.format_exc())



def appendStr100():
    try:
        def get_receiver():
            '''读取收件人列表，以{'公司1': ['邮箱1', '邮箱2'], '公司2': ['邮箱2']}的字典形态返回'''
            receiver_dict = {}
            with open('邮件收件人.txt', 'r', encoding='UTF-8') as contacts_file:
                for a_contact in contacts_file:
                    temp_address_list = []
                    a_contact_list = a_contact.split(',')
                    name = a_contact_list[0].strip()
                    for temp_address in a_contact_list[1:]:
                        temp_address_list.append(temp_address.strip())
                    receiver_dict[name] = temp_address_list
            return receiver_dict

        def read_body(filename):
            '''导入邮件正文的内容'''
            with open(filename, 'r',encoding='UTF-8') as body_file:
                body_file_content = body_file.read()
            return body_file_content

        def put_attachment(file_name, msg):
            '''添加客户附件'''
            part = MIMEBase('attachment', 'octet-stream')
            file_route = attach_file + '\\' + file_name
            part.set_payload(open(file_route, 'rb').read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment',filename="%s" % make_header([(file_name, 'UTF-8')]).encode('UTF-8'))  # 显示中文附件的话选这个
            msg.attach(part)

        # 发件人邮箱和密码
        try:
            # #基因邮箱
            MY_ADDRESS = 'jiabin.zheng@hgt.cn'
            myPass = '8888888'
            server = SMTP('smtp.partner.outlook.cn')

            server.starttls()
            server.login(MY_ADDRESS, myPass)
            path_this_file = os.path.abspath('.') + "\\"

            # 获取邮件正文
         
            email_body = read_body('邮件正文.txt')
            # print(email_body)
            print('>>>获取邮件正文成功！')

            receiver_dict = get_receiver()
            print('>>>获取收件人列表成功！')

            # 获取附件列表
            # attach_file = path_this_file + '群发附件'
            attach_file = path_this_file
            attach_list = os.listdir(attach_file)
            print(attach_list)
            print('>>>获取附件列表成功！')

            for key, value in receiver_dict.items():
                msg = MIMEMultipart()
                msg['From'] = MY_ADDRESS
                receivers = ','.join(value)
                msg['To'] = receivers
                msg['Subject'] = key + '对账单'
                # msg['Subject'] = key
                msg.attach(MIMEText(email_body))  # 邮件正文
                temp_pic_list = []
                for pic in attach_list:
                    if key in pic:
                        put_attachment(pic, msg)
                        temp_pic_list.append(pic)
                if temp_pic_list:
                    server.send_message(msg)
                    print('>>>{}邮件发送成功！'.format(key))
                else:
                    print('>>>{}因无附件，没有发送！'.format(key))
            server.quit()
            print('>>>所有邮件均已发送成功！')
            tkinter.messagebox.showinfo("运行结果", ">>>所有邮件均已发送成功！");
        except:
            traceback.print_exc()


    except Exception as error:

        tm.showerror(title="煎饼提示前方路堵",

                     message="请检查提交源文件是否正确 '" + str(error) + "'.",

                     detail=traceback.format_exc())

appendBtn=Button(a,text="1.点击整理对账单",width=40,height=1,command=appendStrJY01);
appendBtn.pack();
appendBtn=Button(a,text="2.点击整理未开票",width=40,height=1,command=appendStrJY02);
appendBtn.pack();
appendBtn=Button(a,text="3.点击开始发送邮件",width=40,height=1,command=appendStr100);
appendBtn.pack();

tiw,mainloop();
