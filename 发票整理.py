# -*- coding:utf-8 -*-
from tkinter import *;
import pandas as pd
import tkinter.messagebox as tm
import tkinter.filedialog
import datetime
import traceback
#import xlrd



tiw=Tk("煎饼滴小程序0.3");
tiw.title();
tiw.geometry("150x110");
menubar=Menu(tiw)
content=[['煎饼提示:选择功能后请按照提示选择源文件！']]
Main=['煎饼认证0628']
for i in range(len(Main)):
    #新建一个空的菜单,将menubar的menu属性指定为filemenu，即filemenu为menubar的下拉菜单
    filemenu = Menu(menubar, tearoff=0)
    for k in content[i]:
        filemenu.add_command(label = k)
    menubar.add_cascade(label=Main[i], menu=filemenu)

# 将root的menu属性设置为M
tiw['menu'] = menubar
#tiw.mainloop()

a = LabelFrame(tiw, height=22, width=50, text='认证整理功能')
a.pack(side='left', anchor='ne')

def appendStr1():
 try:
    tkinter.messagebox.showinfo("提醒", "请选择平台未认证明细（平台多份的明细请先拼到一个表）");
    df = pd.read_excel(tkinter.filedialog.askopenfilename(),converters = {u'发票号码':str,u'发票代码':str});

    df["发票号码"].astype("int64")

    tkinter.messagebox.showinfo("提醒", "请选择公司本月整理需要认证的明细-包含票号和发票代码和日期");
    #df2 = df1.drop(df1.index[[[[[0, 1, 2, 3,4]]]]], axis=0);#删1-4行
    df2 = pd.read_excel(tkinter.filedialog.askopenfilename(),converters = {u'发票号码':str,u'发票代码':str});
    df2["发票号码"].astype("int64")
    df2['是否勾选(是/否)']='是'

    #开始组合11-22

    df3 = pd.merge(df, df2, how='left', on=['发票号码','发票代码','开票日期']);  # 完全相同合并，忽略没有的货品ID(没有how)
    df4 = df3[df3["是否勾选(是/否)_y"] == "是"]
    df5 = df4.rename(columns={'是否勾选(是/否)_y': '是否勾选(是/否)'});
    df6 = df5.drop(columns={'是否勾选(是/否)_x'})
    df6["发票号码"].astype("int64")

    df6["发票代码1"] = df6["发票代码"]
    df6["发票号码1"] = df6["发票号码"]
    df6["开票日期1"] = df6["开票日期"]
    df6["税额1"] = df6["税额"]
    df6["有效抵扣税额1"] = df6["有效抵扣税额"]
    df6["销方名称1"] = df6["销方名称"]
    df6["销方税号1"] = df6["销方税号"]
    df6["金额1"] = df6["金额"]
    df6["用途1"] = df6["用途"]

    df7=df6.drop(["发票代码", "发票号码", "开票日期","税额","有效抵扣税额", "销方名称","销方税号","金额","用途","发票类型","管理状态"], axis=1)  # 删列
    df8=df7.rename(columns={'发票代码1': '发票代码','发票号码1': '发票号码','开票日期1': '开票日期','税额1': '税额','有效抵扣税额1': '有效抵扣税额','销方名称1': '销方名称','销方税号1': '销方税号','金额1': '金额','用途1': '用途'});

    #########

    df300 = pd.merge(df2, df, how='left', on=['发票号码']);
    df301 = df300[df300["是否勾选(是/否)_y"] != "否"]

    df302 = df301.drop(["是否勾选(是/否)_x", "是否勾选(是/否)_y", "发票代码", "开票日期", "税额", "有效抵扣税额", "销方名称", "销方税号", "金额", "用途"],
                       axis=1)  # 删列

    print(df8)

    df8.to_excel("本次认证整理文件" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                 index=False)  # 自动输出
    df302.to_excel("本次认证失败明细" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                   index=False)  # 自动输出

    tkinter.messagebox.showinfo("运行结果","需认证整理成功!");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())

def appendStr2():
 try:
    tkinter.messagebox.showinfo("提醒", "请选择平台未认证明细（平台多份的明细请先拼到一个表）");
    df = pd.read_excel(tkinter.filedialog.askopenfilename(),converters = {u'发票号码':str,u'发票代码':str});

    df["发票号码"].astype("int64")

    tkinter.messagebox.showinfo("提醒", "请选择公司本月整理需要认证的明细-包含票号和发票代码");
    #df2 = df1.drop(df1.index[[[[[0, 1, 2, 3,4]]]]], axis=0);#删1-4行
    df2 = pd.read_excel(tkinter.filedialog.askopenfilename(),converters = {u'发票号码':str,u'发票代码':str});
    df2["发票号码"].astype("int64")
    df2['是否勾选(是/否)']='是'

    #开始组合11-22

    df3 = pd.merge(df, df2, how='left', on=['发票号码','发票代码']);  # 完全相同合并，忽略没有的货品ID(没有how)
    df4 = df3[df3["是否勾选(是/否)_y"] == "是"]
    df5 = df4.rename(columns={'是否勾选(是/否)_y': '是否勾选(是/否)'});
    df6 = df5.drop(columns={'是否勾选(是/否)_x'})
    df6["发票号码"].astype("int64")

    df6["发票代码1"] = df6["发票代码"]
    df6["发票号码1"] = df6["发票号码"]
    df6["开票日期1"] = df6["开票日期"]
    df6["税额1"] = df6["税额"]
    df6["有效抵扣税额1"] = df6["有效抵扣税额"]
    df6["销方名称1"] = df6["销方名称"]
    df6["销方税号1"] = df6["销方税号"]
    df6["金额1"] = df6["金额"]
    df6["用途1"] = df6["用途"]

    df7=df6.drop(["发票代码", "发票号码", "开票日期","税额","有效抵扣税额", "销方名称","销方税号","金额","用途","发票类型","管理状态"], axis=1)  # 删列
    df8=df7.rename(columns={'发票代码1': '发票代码','发票号码1': '发票号码','开票日期1': '开票日期','税额1': '税额','有效抵扣税额1': '有效抵扣税额','销方名称1': '销方名称','销方税号1': '销方税号','金额1': '金额','用途1': '用途'});

    #########

    df300 = pd.merge(df2, df, how='left', on=['发票号码']);
    df301 = df300[df300["是否勾选(是/否)_y"] != "否"]

    df302 = df301.drop(["是否勾选(是/否)_x", "是否勾选(是/否)_y", "发票代码", "开票日期", "税额", "有效抵扣税额", "销方名称", "销方税号", "金额", "用途"],
                       axis=1)  # 删列

    print(df8)

    df8.to_excel("本次认证整理文件" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                 index=False)  # 自动输出
    df302.to_excel("本次认证失败明细" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                   index=False)  # 自动输出


    tkinter.messagebox.showinfo("运行结果","需认证整理成功!");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())

def appendStr3():
 try:
    tkinter.messagebox.showinfo("提醒", "请选择平台未认证明细（平台多份的明细请先拼到一个表）");
    #df = pd.read_excel(tkinter.filedialog.askopenfilename());
    df = pd.read_excel(tkinter.filedialog.askopenfilename(), converters={u'发票号码': str, u'发票代码': str});
    df["发票号码"].astype("int64")

    tkinter.messagebox.showinfo("提醒", "请选择公司本月整理需要认证的明细-包含票号");
    #df2 = df1.drop(df1.index[[[[[0, 1, 2, 3,4]]]]], axis=0);#删1-4行
    df2 = pd.read_excel(tkinter.filedialog.askopenfilename(),converters = {u'发票号码':str,u'发票代码':str});
    df2["发票号码"].astype("int64")
    df2['是否勾选(是/否)']='是'

    #开始组合11-22

    df3 = pd.merge(df, df2, how='left', on=['发票号码']);  # 完全相同合并，忽略没有的货品ID(没有how)


    df4 = df3[df3["是否勾选(是/否)_y"] == "是"]
    df5 = df4.rename(columns={'是否勾选(是/否)_y': '是否勾选(是/否)'});
    df6 = df5.drop(columns={'是否勾选(是/否)_x'})
    df6["发票号码"].astype("int64")

    df6["发票代码1"] = df6["发票代码"]
    df6["发票号码1"] = df6["发票号码"]
    df6["开票日期1"] = df6["开票日期"]
    df6["税额1"] = df6["税额"]
    df6["有效抵扣税额1"] = df6["有效抵扣税额"]
    df6["销方名称1"] = df6["销方名称"]
    df6["销方税号1"] = df6["销方税号"]
    df6["金额1"] = df6["金额"]
    df6["用途1"] = df6["用途"]

    df7=df6.drop(["发票代码", "发票号码", "开票日期","税额","有效抵扣税额", "销方名称","销方税号","金额","用途","发票类型","管理状态"], axis=1)  # 删列
    df8=df7.rename(columns={'发票代码1': '发票代码','发票号码1': '发票号码','开票日期1': '开票日期','税额1': '税额','有效抵扣税额1': '有效抵扣税额','销方名称1': '销方名称','销方税号1': '销方税号','金额1': '金额','用途1': '用途'});
    df81=df8.drop(["Unnamed: 1","Unnamed: 2","Unnamed: 3","Unnamed: 4","Unnamed: 5","Unnamed: 6","Unnamed: 7"],axis=1)
    #########

    df300 = pd.merge(df2, df, how='left', on=['发票号码']);
    df301=df300[df300["是否勾选(是/否)_y"] != "否"]

    df302=df301.drop(["是否勾选(是/否)_x", "是否勾选(是/否)_y", "发票代码", "开票日期","税额","有效抵扣税额","销方名称","销方税号","金额","用途","Unnamed: 1","Unnamed: 2","Unnamed: 3","Unnamed: 4","Unnamed: 5","Unnamed: 6","Unnamed: 7"], axis=1)  # 删列

    print(df8)

    df81.to_excel("本次认证整理文件" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                    index=False)  # 自动输出
    df302.to_excel("本次认证失败明细" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                 index=False)  # 自动输出


    tkinter.messagebox.showinfo("运行结果","需认证整理成功!");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())

appendBtn=Button(a,text="发票代码+号码+日期",width=22,height=1,command=appendStr1);
appendBtn.pack();
appendBtn=Button(a,text="发票代码+号码",width=22,height=1,command=appendStr2);
appendBtn.pack();
appendBtn=Button(a,text="发票号码",width=22,height=1,command=appendStr3);
appendBtn.pack();
tiw,mainloop();
