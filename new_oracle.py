# -*- coding:utf-8 -*-
import tkinter
from tkinter import *;
import tkinter.messagebox as tm
import traceback
import tkinter.filedialog
import pandas as pd
import datetime


import urllib
import urllib.parse
import urllib.request
import base64
import json
import os
# import cx_Oracle as oracle

import xlwt

import time


def gettime():
   # 获取当前时间并转为字符串
   timestr = time.strftime("%D-%H:%M:%S")
   # 重新设置标签文本
   lb.configure(text="当前时间："+ timestr +"  注意休息 请勿熬夜")
   # 每隔一秒调用函数gettime自身获取时间
   root.after(1000, gettime)

root = Tk()
root.title("煎饼滴百宝箱2.62-20201207");
root.geometry("510x300");
root.resizable(0,0)   #禁止调整窗口大小
#######1118
# tk.Label(root, text='运行进度:', ).place(x=50, y=60)
# canvas = tk.Canvas(root, width=300, height=20, bg="white")
# canvas.place(x=400, y=600)
#######1118


photo =PhotoImage(file="LOGO.gif")
imglabel=Label(root,image=photo)

# 设置字体大小颜色
lb = tkinter.Label(root, text='', fg='black',font=("黑体", 10))
lb.pack(side=BOTTOM)  ##这里的side可以赋值为LEFT  RTGHT TOP  BOTTOM
gettime()

imglabel.pack()



def callback():
    print('调用')
menubar = Menu(root)




def appendStr():
 try:
    # tkinter.messagebox.showinfo("提醒", "请选择开票明细表");
    # df = pd.read_excel("d:/data.xlsx", sheet_name="sheet1");  #
    df = pd.read_excel(tkinter.filedialog.askopenfilename());
    df1 = df.drop(df.columns[[[[[[[[[[[0, 3, 4, 5, 7, 8, 10, 11, 12, 13, 17]]]]]]]]]]], axis=1);
    df2 = df1.drop(df1.index[[[[0, 1, 2, 3]]]], axis=0);
    df3 = df2[df2["Unnamed: 9"] != "小计"];
    df4 = df3[df3["Unnamed: 9"] != "商品名称"];
    df5 = df4.dropna(how="all");
    df6 = df5.fillna(method='pad');
    df6["Unnamed: 14"] = df6["Unnamed: 14"].astype("float64");  # 改变格式
    df6["Unnamed: 16"] = df6["Unnamed: 16"].astype("float64");

    df10 = df6

    #df10 = pd.read_excel("d:/data4.xlsx", sheet_name="测试1")
    df10["Unnamed: 14"] = df10["Unnamed: 14"].astype("float64");
    df10["Unnamed: 16"] = df10["Unnamed: 16"].astype("float64");
    df11 = df10.groupby("Unnamed: 2")["Unnamed: 14","Unnamed: 16"].sum() ;
    df12 = df11['发票合计'] = df11.apply(lambda x: x.sum(), axis=1);

    print(df12)

    #tkinter.filedialog.asksaveasfile(mode='w',
     #   defaultextension='.txt',

    #df12.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb',defaultextension='*.xlsx',));#指定位置另存为630

    df12.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"));

    tkinter.messagebox.showinfo("运行结果","客户汇总成功!");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())

def appendStr1():#税率
 try:
    tkinter.messagebox.showinfo("提醒", "请选择开票明细表");
    #df = pd.read_excel("d:/data.xlsx", sheet_name="sheet1");  #
    df = pd.read_excel(tkinter.filedialog.askopenfilename());
    df1 = df.drop(df.columns[[[[[[[[[[[0,3,4,5,7,8,10,11,12,13,17]]]]]]]]]]], axis=1) ;
    df2 = df1.drop(df1.index[[[[0, 1, 2, 3]]]], axis = 0);
    df3 = df2[df2["Unnamed: 9"] != "小计"];
    df4 = df3[df3["Unnamed: 9"] != "商品名称"];
    df5 = df4.dropna(how="all");
    df6 = df5.fillna(method='pad');
    df6["Unnamed: 14"]=df6["Unnamed: 14"].astype("float64");#改变格式
    df6["Unnamed: 16"]=df6["Unnamed: 16"].astype("float64");


    print(df6)
    #df6.to_excel(excel_writer="d:/data4.xlsx",
     #            sheet_name="测试1",
     #            index = False);

    df6.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                                   "Microsoft Excel 97-20003 文件",
                                                                                   "*.xls")],
                                                                    defaultextension=".xlsx"));
    tkinter.messagebox.showinfo("运行结果","销售明细含税率导出成功！");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())




def appendStr2():
 try:
    tkinter.messagebox.showinfo("提醒", "请选择开票明细表");
    # df = pd.read_excel("d:/data.xlsx", sheet_name="sheet1");  #
    df = pd.read_excel(tkinter.filedialog.askopenfilename());
    df1 = df.drop(df.columns[[[[[[[[[[[0, 3, 4, 5, 7, 8, 10, 11, 12, 13, 17]]]]]]]]]]], axis=1);
    df2 = df1.drop(df1.index[[[[0, 1, 2, 3]]]], axis=0);
    df3 = df2[df2["Unnamed: 9"] != "小计"];
    df4 = df3[df3["Unnamed: 9"] != "商品名称"];
    df5 = df4.dropna(how="all");
    df6 = df5.fillna(method='pad');
    df6["Unnamed: 14"] = df6["Unnamed: 14"].astype("float64");  # 改变格式
    df6["Unnamed: 16"] = df6["Unnamed: 16"].astype("float64");


    df13 = df6


    #df13 = pd.read_excel("d:/data4.xlsx", sheet_name="测试1")  #
    df13["Unnamed: 14"] = df13["Unnamed: 14"].astype("float64");
    df13["Unnamed: 16"] = df13["Unnamed: 16"].astype("float64");
    df14 = df13.groupby("Unnamed: 1")["Unnamed: 14", "Unnamed: 16"].sum();
    df15 = df14['发票合计'] = df14.apply(lambda x: x.sum(), axis=1);

    print(df15)


    #df15.to_excel(excel_writer="d:/发票汇总.xlsx",
    #              sheet_name="按发票汇总");

    #df15.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630
    df15.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"));
    tkinter.messagebox.showinfo("运行结果", "发票汇总成功!");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())

def appendStr3():#客户分列单体账龄
 try:
    tkinter.messagebox.showinfo("提醒", "请选择金蝶单体公司账龄表");
    df16 = pd.read_excel(tkinter.filedialog.askopenfilename());
    df18 = df16['客户'].str.split('-| ',expand=True);
    print(df18)
    df19 = pd.merge(df18, df16, right_index=True, left_index=True);
    df20 = df19.drop(df19.columns[[[[8,10,12,13]]]], axis=1)
    df22 = df20.dropna();

    #df22.to_excel(excel_writer="d:/分列前端数据.xlsx",
      #                     sheet_name="测试1",
       #                   index=False);


    df22['客户'] = df22[7]
    df22['部门'] = df22[0]
    df22['片区'] = df22[1]
    df22['客户编码'] = df22[3]

    print(df22)
    df23 = df22.groupby(["部门","片区","客户","客户编码"])[ "过期", "Unnamed: 7", "Unnamed: 8", "Unnamed: 9", "Unnamed: 10", "Unnamed: 11"].sum();
    c_df = pd.DataFrame(df23)
    c_df.reset_index(inplace=True)  # 取消合并

    df24 = df23.rename(columns={'过期': '1-30','Unnamed: 7':'31-60','Unnamed: 8':'61-90','Unnamed: 9':'91-120',
                                 'Unnamed: 10':'121-150','Unnamed: 11':'151-'});

    df24["部门"].replace("一部", "01部", inplace=True)
    df24["部门"].replace("二部", "02部", inplace=True)
    df24["部门"].replace("三部", "03部", inplace=True)
    df24["部门"].replace("四1部", "04部1", inplace=True)
    df24["部门"].replace("四2部", "04部2", inplace=True)
    df24["部门"].replace("五1部", "05部1", inplace=True)
    df24["部门"].replace("五2部", "05部2", inplace=True)
    df24["部门"].replace("六部", "06部", inplace=True)
    df24["部门"].replace("七部", "07部", inplace=True)
    df24["部门"].replace("八部", "08部", inplace=True)
    df24["部门"].replace("九部", "09部", inplace=True)
    df24["部门"].replace("十部", "10部", inplace=True)
    df24["部门"].replace("十一部", "11部", inplace=True)
    df24["部门"].replace("十二部", "12部", inplace=True)
    df24["片区"].replace("", "无", inplace=True)
    df24 = df24[df24["片区"] != "无"]
    df24["余额"] = df24["1-30"] + df24["31-60"] + df24["61-90"] + df24["91-120"] + df24["121-150"] + df24["151-"]

    #df24.to_excel(excel_writer="d:/分列前端数据.xlsx",
                              #            sheet_name="测试1",
                              #           index=False);



    df25 = df24.sort_values(by=['部门','片区'], axis=0, ascending=True)
    print(df25)

    df26 = df25.groupby(["部门", "片区", "客户编码","客户","余额"])[
        "1-30", "31-60", "61-90", "91-120", "121-150", "151-"].sum();
    c_df = pd.DataFrame(df26)
    c_df.reset_index(inplace=True)  # 取消合并



   # df26.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630
    df26.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"));
    tkinter.messagebox.showinfo("运行结果","分列并导出成功！");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())


def appendStr5():#计算发出商品
 try:
    tkinter.messagebox.showinfo("提醒", "请先选择英克进销存源文件");
    #df23 = pd.read_excel("d:/进销存.xlsx", sheet_name="查询存货账汇总");#630换路径
    df23 = pd.read_excel(tkinter.filedialog.askopenfilename());

    df24 = df23.groupby("货品ID")["本期金额","本期数量","进货金额","进货数量","上期金额","上期数量","调整金额","销售发票数量","销售发票金额"].sum();

    df24['本期成本单价'] = (df24['本期金额']+df24['上期金额']+df24['进货金额']+df24["销售发票金额"]-df24['调整金额'])/(df24['本期数量']+df24['上期数量']+df24['进货数量']+df24["销售发票数量"]) # 后面增加一列相除单价
    print(df24);

    #df241 = pd.read_excel("d:/英克本期成本单价.xlsx", sheet_name="货品成本单价");

    #df25 = pd.read_excel("d:/目前未开票.xlsx", sheet_name="销售开发票");#630换路径
    tkinter.messagebox.showinfo("提醒", "请先选择英克未开票源文件");
    df25 = pd.read_excel(tkinter.filedialog.askopenfilename());

    #df26 = df25.groupby("货品ID")["未开票数量"].sum();#0702暂时修改为未发票数量
    #修改列名1201
    df251 = df25.rename(columns={'未结算数量': '未开票数量'});
    df26 = df251.groupby("货品ID")["未开票数量"].sum();

    #df26 = pd.merge(df23, df25, right_index=True, left_index=True);
    df27 = pd.merge(df26, df24,how='left',on=['货品ID'] );#完全相同合并，忽略没有的货品ID(没有how)
    #df27 = pd.merge(df26, df241);

    #df27['本期发出商品金额'] = df27['本期成本单价'] * df27['未开票数量']#0702暂时修改为未发票数量
    df27['本期发出商品金额'] = df27['本期成本单价'] * df27['未开票数量']
    print(df27);

    #df27.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='*.xlsx', ));  # 指定位置另存为630
    df27.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"));


    tkinter.messagebox.showinfo("运行结果", "发出商品导出成功！");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())


def appendStr6():#统计未开票客户
 try:

    tkinter.messagebox.showinfo("提醒", "请先选择英克未开票源文件");

    df31 = pd.read_excel(tkinter.filedialog.askopenfilename());#630自主选择路径

    df32 = df31[df31["未开票金额"]!=0];

    df32['发票月份'] = df32['业务日期'].map(lambda x: 100*x.year + x.month);#转换日期

    df33 = df32.groupby(['客户','发票月份'])["未开票金额"].sum();

    c_df = pd.DataFrame(df33)
    c_df.reset_index(inplace=True)

    print(c_df)

    df35 = pd.pivot_table(c_df, index="客户",columns="发票月份",values="未开票金额");#列换行

    df35['合计'] = df35.apply(lambda x: x.sum(), axis=1)


    tkinter.messagebox.showinfo("提醒", "请先选择英克客户分类源文件");
    df351 = pd.read_excel(tkinter.filedialog.askopenfilename());#630修改为手工读路径

    df352 = df351.rename(columns={'客户名称': '客户'}) #把原来的 客户名称 命名为 客户
    # 查看读取数据内容
    print(df352)

    # 查看是否有重复行
    re_row = df352.duplicated()
    print(re_row)

    # 查看去除重复行的数据
    no_re_row = df352.drop_duplicates()
    print(no_re_row)

    # 查看基于[物品]列去除重复行的数据f
    wp = df352.drop_duplicates(['客户'])
    print(wp)



    df3521 = pd.DataFrame([["艾康生物技术（杭州）有限公司","调拨","试剂调拨部"],
                        ["艾美卫信生物药业（浙江）有限公司","奉化宁海象山","诊断二部"]
                            ]
                       ,columns = ["客户","诊断客户地区","诊断客户部门"]);

    #建立内部客户片区中间表6-27

    print(df3521)

   # excelFile = r'D:\5.xlsx'
   # df3522 = pd.DataFrame(pd.read_excel(excelFile))
   # print(df3522)  #直接读取表客户片区放弃6-27

   # df353 = pd.read_excel("E:/5.xlsx", sheet_name="去重客户区域")


    df36 = pd.merge(df35, wp, how='left', on=['客户'])#组合

    #df37 = df36.drop(df36.columns[[[[7,8,9,12]]]], axis = 1)

    df37 = df36.drop(["公司名称","客户ID ","药品部门","财务编码"],axis = 1)#删列

    print(df37)

    #df37.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='*.xlsx', ));  # 指定位置另存为630

    df37.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"));
    tkinter.messagebox.showinfo("运行结果", "客户未开票导出成功！");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())


def appendStr8():#辅助总账分列(黄)
 try:
    tkinter.messagebox.showinfo("提醒", "请选择金蝶辅助总账表");
    df50 = pd.read_excel(tkinter.filedialog.askopenfilename());
    df51 = df50.drop(df50.index[0], axis=0);
    df52 = df51[df51["Unnamed: 1"]!="小计"]

    df53 = df52['辅助账:'].str.split('-| ',expand=True);

    df54 = pd.merge(df53, df50, right_index=True, left_index=True);

    df55 = df54.drop(columns=["辅助账:"])

    print(df55)
   # df55.to_excel(excel_writer="e:/辅助账处理.xlsx",
   #              sheet_name="处理",
   #              index=False);

   # df56 = pd.read_excel("e:/辅助账处理.xlsx", sheet_name="处理");
    #df57= df56.dropna();#去除空白
    df58 = df55.rename(columns={'7': '客户名称', '0': '客户编码', 'Unnamed: 3': '会计期间', 'Unnamed: 5': '期初余额', '期间:2019年第6期': '期初方向',
                 'Unnamed: 6': '借方', 'Unnamed: 7': '贷方','币别:人民币':'本年累计借方','Unnamed: 9':'本年累计贷方','Unnamed: 1':'科目编码',
                                '核算项目:客户':'会计科目', '浙江海尔施医疗设备有限公司':'期末方向','Unnamed: 11':'期末余额'});


    print(df58)

    #df58.columns = ['部门', '片区', '片区编码','客户编码',"医药区域","医药区域2","账龄", "客户"]

   # df58.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630

    df58.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"));
    tkinter.messagebox.showinfo("运行结果", "应收账款辅助总账表导出成功！");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())

def appendStr9():#集团客户账龄分列
 try:
    tkinter.messagebox.showinfo("提醒", "请选择金蝶集团公司账龄表");
    df59 = pd.read_excel(tkinter.filedialog.askopenfilename());
    #df60 = df59.drop(df59.index[0], axis=0);
    df61 = df59[df59["公司"] != "海尔施集团"]
    df62 = df61['客户'].str.split('-| ', expand=True);

    print(df62)

    df63 = pd.merge(df62, df59, right_index=True, left_index=True);
    df64 = df63.drop(["客户","项目"], axis=1)  # 删列
    df641 = df64.dropna();
    df68 = df641[df641["过期"] != "1- 30"]

    #df68 = df65.rename( columns={'过期': '1-30', 'Unnamed: 8': '31-60', 'Unnamed: 9': '61-90', 'Unnamed: 10': '91-120',
     #                  'Unnamed: 11': '121-150', 'Unnamed: 12': '151-'});

    df68["客户"] = df68[7]
    df68["部门"] = df68[0]
    df68["片区"] = df68[1]
    df68['客户编码'] = df68[3]
    df68['片区编码'] = df68[2]

    print(68)

    df70 = df68.groupby(["部门","片区","片区编码","客户","客户编码","公司"])[ "过期", "Unnamed: 8", "Unnamed: 9", "Unnamed: 10", "Unnamed: 11", "Unnamed: 12"].sum();

    #df70 = df68.groupby(["客户编码", "公司"])["余额", "1-30", "31-60", "61-90", "91-120", "121-150", "151-"].sum();
    #仅需在groupby命令后添加 as_index=False的参数即可，表明不希望标签值作为索引：703无效


    # df70 = pd.pivot_table(df69,index=["部门","片区","客户编码","客户", "公司"],valuse=["余额"])错误，不采用


    c_df = pd.DataFrame(df70)
    c_df.reset_index(inplace=True)#取消合并

    df71 = df70.rename(columns={'过期': '1-30', 'Unnamed: 8': '31-60', 'Unnamed: 9': '61-90', 'Unnamed: 10': '91-120',
                                'Unnamed: 11': '121-150', 'Unnamed: 12': '151-'});

    #df70['客户编码']= df68[3]

    #df71 = pd.read_excel("E:/集团分列账龄暂存数据.xlsx", sheet_name="1")
    #print(df71)


    df71["部门"].replace("一部","01部",inplace=True)
    df71["部门"].replace("二部", "02部", inplace=True)
    df71["部门"].replace("三部", "03部", inplace=True)
    df71["部门"].replace("四部", "04部", inplace=True)
    df71["部门"].replace("五部", "05部", inplace=True)
    df71["部门"].replace("六部", "06部", inplace=True)
    df71["部门"].replace("七部", "07部", inplace=True)
    df71["部门"].replace("八部", "08部", inplace=True)
    df71["部门"].replace("九部", "09部", inplace=True)
    df71["部门"].replace("十部", "10部", inplace=True)
    df71["部门"].replace("十一部", "11部", inplace=True)
    df71["部门"].replace("十二部", "12部", inplace=True)
    df71["片区"].replace("", "无", inplace=True)
    df72=df71[df71["片区"] != "无"]
    df72["余额"]=df72["1-30"]+df72["31-60"]+df72["61-90"]+df72["91-120"]+df72["121-150"]+df72["151-"]
   #以下增加序列号码唯一性排序
    df72["序列号码"]=df72["片区"]
    print(df72)
    #以下是区域序列划分表
    df72["序列号码"].replace("温州1", "0101", inplace=True)
    df72["序列号码"].replace("温州2", "0102", inplace=True)
    df72["序列号码"].replace("台州1", "0103", inplace=True)
    df72["序列号码"].replace("台州2", "0104", inplace=True)
    df72["序列号码"].replace("丽水", "0105", inplace=True)
    df72["序列号码"].replace("宁波", "0201", inplace=True)
    df72["序列号码"].replace("舟山北仑", "0202", inplace=True)
    df72["序列号码"].replace("北三县", "0203", inplace=True)
    df72["序列号码"].replace("南三县", "0204", inplace=True)
    df72["序列号码"].replace("杭州姜立民", "0301", inplace=True)
    df72["序列号码"].replace("杭州周海波", "030201", inplace=True)
    df72["序列号码"].replace("杭州沈剑芳", "030202", inplace=True)
    df72["序列号码"].replace("杭州石亚国", "030203", inplace=True)
    df72["序列号码"].replace("嘉兴阮芳", "030301", inplace=True)
    df72["序列号码"].replace("湖州陈荣斌", "030302", inplace=True)
    df72["序列号码"].replace("晋江运城郑良", "0304", inplace=True)
    df72["序列号码"].replace("南京高跃", "0401", inplace=True)
    df72["序列号码"].replace("南京阮建锋", "0402", inplace=True)
    df72["序列号码"].replace("南京刘纪彬", "0403", inplace=True)
    df72["序列号码"].replace("南京陈豪", "0404", inplace=True)
    df72["序列号码"].replace("南通朱一亦", "0501", inplace=True)
    df72["序列号码"].replace("南通王峥骅", "0502", inplace=True)
    df72["序列号码"].replace("盐城", "0503", inplace=True)
    df72["序列号码"].replace("连云港", "0504", inplace=True)
    df72["序列号码"].replace("上海1", "0601", inplace=True)
    df72["序列号码"].replace("上海2", "0602", inplace=True)
    df72["序列号码"].replace("苏州", "0701", inplace=True)
    df72["序列号码"].replace("苏州市郊", "0702", inplace=True)
    df72["序列号码"].replace("扬泰1", "0801", inplace=True)
    df72["序列号码"].replace("扬泰2", "0802", inplace=True)
    df72["序列号码"].replace("扬泰3", "0803", inplace=True)
    df72["序列号码"].replace("徐州于博", "090101", inplace=True)  #
    df72["序列号码"].replace("徐州张浩", "090102", inplace=True)  #
    df72["序列号码"].replace("宿迁", "0902", inplace=True)
    df72["序列号码"].replace("淮安", "0903", inplace=True)
    df72["序列号码"].replace("常州", "1001", inplace=True)
    df72["序列号码"].replace("镇江", "1002", inplace=True)
    df72["序列号码"].replace("常镇", "1003", inplace=True)
    df72["序列号码"].replace("绍兴龚群波", "1101", inplace=True)
    df72["序列号码"].replace("衢州", "110201", inplace=True)  #
    df72["序列号码"].replace("金华", "110202", inplace=True)  #
    df72["序列号码"].replace("无锡1", "1201", inplace=True)
    df72["序列号码"].replace("无锡2", "1202", inplace=True)


    #df72.to_excel(excel_writer="e:/集团账龄数据1.xlsx",
        #           sheet_name="处理",
        #           index=False);
    #新建一张表用于拼接 703

    df73 = df72.groupby(["部门"],as_index=False)[ "余额", "1-30", "31-60", "61-90", "91-120", "121-150","151-"].sum();

    df73["客户"]=df73["部门"]
    df73["序列号码"]= df73["部门"]

    df73["客户"].replace("01部", "01部合计", inplace=True)
    df73["客户"].replace("02部", "02部合计", inplace=True)
    df73["客户"].replace("03部", "03部合计", inplace=True)
    df73["客户"].replace("04部", "04部合计", inplace=True)
    df73["客户"].replace("05部", "05部合计", inplace=True)
    df73["客户"].replace("06部", "06部合计", inplace=True)
    df73["客户"].replace("07部", "07部合计", inplace=True)
    df73["客户"].replace("08部", "08部合计", inplace=True)
    df73["客户"].replace("09部", "09部合计", inplace=True)
    df73["客户"].replace("10部", "10部合计", inplace=True)
    df73["客户"].replace("11部", "11部合计", inplace=True)
    df73["客户"].replace("12部", "12部合计", inplace=True)


    df73["序列号码"].replace("01部", "0199", inplace=True)
    df73["序列号码"].replace("02部", "0299", inplace=True)
    df73["序列号码"].replace("03部", "0399", inplace=True)
    df73["序列号码"].replace("04部", "0499", inplace=True)
    df73["序列号码"].replace("05部", "0599", inplace=True)
    df73["序列号码"].replace("06部", "0699", inplace=True)
    df73["序列号码"].replace("07部", "0799", inplace=True)
    df73["序列号码"].replace("08部", "0899", inplace=True)
    df73["序列号码"].replace("09部", "0999", inplace=True)
    df73["序列号码"].replace("10部", "1099", inplace=True)
    df73["序列号码"].replace("11部", "1199", inplace=True)
    df73["序列号码"].replace("12部", "1299", inplace=True)

    #df73.loc["总计"] = df73.apply(lambda x: x.sum())
    #df73["序列号码"].replace("12部", "1299", inplace=True)


    #df73.to_excel(excel_writer="e:/集团账龄数据.xlsx",
     #                  sheet_name="处理",
      #                 index=False);

    df731 = df72.groupby(["片区"],as_index=False)["余额", "1-30", "31-60", "61-90", "91-120", "121-150", "151-"].sum();



    df731["序列号码"] = df731["片区"]
    df731["客户"] = df731["片区"]


    df731["序列号码"].replace("温州1", "010199", inplace=True)
    df731["序列号码"].replace("温州2", "010299", inplace=True)
    df731["序列号码"].replace("台州1", "010399", inplace=True)
    df731["序列号码"].replace("台州2", "010499", inplace=True)
    df731["序列号码"].replace("丽水", "010599", inplace=True)
    df731["序列号码"].replace("宁波", "020199", inplace=True)
    df731["序列号码"].replace("舟山北仑", "020299", inplace=True)
    df731["序列号码"].replace("北三县", "020399", inplace=True)
    df731["序列号码"].replace("南三县", "020499", inplace=True)
    df731["序列号码"].replace("杭州姜立民", "030199", inplace=True)
    df731["序列号码"].replace("杭州周海波", "03020199", inplace=True)#
    df731["序列号码"].replace("杭州沈剑芳", "03020299", inplace=True)#
    df731["序列号码"].replace("杭州石亚国", "03020399", inplace=True)#
    df731["序列号码"].replace("嘉兴阮芳", "03030199", inplace=True)#
    df731["序列号码"].replace("湖州陈荣斌", "03030299", inplace=True)#
    df731["序列号码"].replace("晋江运城郑良", "030499", inplace=True)
    df731["序列号码"].replace("南京高跃", "040199", inplace=True)
    df731["序列号码"].replace("南京阮建锋", "040299", inplace=True)
    df731["序列号码"].replace("南京刘纪彬", "040399", inplace=True)
    df731["序列号码"].replace("南京陈豪", "040499", inplace=True)
    df731["序列号码"].replace("南通朱一亦", "050199", inplace=True)
    df731["序列号码"].replace("南通王峥骅", "050299", inplace=True)
    df731["序列号码"].replace("盐城", "050399", inplace=True)
    df731["序列号码"].replace("连云港", "050499", inplace=True)
    df731["序列号码"].replace("上海1", "060199", inplace=True)
    df731["序列号码"].replace("上海2", "060299", inplace=True)
    df731["序列号码"].replace("苏州", "070199", inplace=True)
    df731["序列号码"].replace("苏州市郊", "070299", inplace=True)
    df731["序列号码"].replace("扬泰1", "080199", inplace=True)
    df731["序列号码"].replace("扬泰2", "080299", inplace=True)
    df731["序列号码"].replace("扬泰3", "080399", inplace=True)
    df731["序列号码"].replace("徐州于博", "09010199", inplace=True)#
    df731["序列号码"].replace("徐州张浩", "09010299", inplace=True)#
    df731["序列号码"].replace("宿迁", "090299", inplace=True)
    df731["序列号码"].replace("淮安", "090399", inplace=True)
    df731["序列号码"].replace("常州", "100199", inplace=True)
    df731["序列号码"].replace("镇江", "100299", inplace=True)
    df731["序列号码"].replace("常镇", "100399", inplace=True)
    df731["序列号码"].replace("绍兴龚群波", "110199", inplace=True)
    df731["序列号码"].replace("衢州", "11020199", inplace=True)#
    df731["序列号码"].replace("金华", "11020299", inplace=True)#
    df731["序列号码"].replace("无锡1", "120199", inplace=True)
    df731["序列号码"].replace("无锡2", "120299", inplace=True)


    df731["客户"].replace("温州1", "温州1合计", inplace=True)
    df731["客户"].replace("温州2", "温州2合计", inplace=True)
    df731["客户"].replace("台州1", "台州1合计", inplace=True)
    df731["客户"].replace("台州2", "台州2合计", inplace=True)
    df731["客户"].replace("丽水", "丽水合计", inplace=True)
    df731["客户"].replace("宁波", "宁波合计", inplace=True)
    df731["客户"].replace("舟山北仑", "舟山北仑合计", inplace=True)
    df731["客户"].replace("北三县", "北三县合计", inplace=True)
    df731["客户"].replace("南三县", "南三县合计", inplace=True)
    df731["客户"].replace("杭州姜立民", "杭州姜立民合计", inplace=True)
    df731["客户"].replace("杭州周海波", "杭州周海波合计", inplace=True)
    df731["客户"].replace("杭州沈剑芳", "杭州沈剑芳合计", inplace=True)
    df731["客户"].replace("杭州石亚国", "杭州石亚国合计", inplace=True)
    df731["客户"].replace("嘉兴阮芳", "嘉兴阮芳合计", inplace=True)
    df731["客户"].replace("湖州陈荣斌", "湖州陈荣斌合计", inplace=True)
    df731["客户"].replace("晋江运城郑良", "晋江运城郑良合计", inplace=True)

    df731["客户"].replace("南京高跃", "南京高跃合计", inplace=True)
    df731["客户"].replace("南京阮建锋", "南京阮建锋合计", inplace=True)
    df731["客户"].replace("南京刘纪彬", "南京刘纪彬合计", inplace=True)
    df731["客户"].replace("南京陈豪", "南京陈豪合计", inplace=True)
    df731["客户"].replace("南通朱一亦", "南通朱一亦合计", inplace=True)
    df731["客户"].replace("南通王峥骅", "南通王峥骅合计", inplace=True)
    df731["客户"].replace("盐城", "盐城合计", inplace=True)
    df731["客户"].replace("连云港", "连云港合计", inplace=True)
    df731["客户"].replace("上海1", "上海1合计", inplace=True)
    df731["客户"].replace("上海2", "上海2合计", inplace=True)
    df731["客户"].replace("苏州", "苏州合计", inplace=True)
    df731["客户"].replace("苏州市郊", "苏州市郊合计", inplace=True)
    df731["客户"].replace("扬泰1", "扬泰1合计", inplace=True)
    df731["客户"].replace("扬泰2", "扬泰2合计", inplace=True)
    df731["客户"].replace("扬泰3", "杨泰3合计", inplace=True)
    df731["客户"].replace("徐州于博", "徐州于博合计", inplace=True)  #
    df731["客户"].replace("徐州张浩", "徐州张浩合计", inplace=True)  #
    df731["客户"].replace("宿迁", "宿迁合计", inplace=True)
    df731["客户"].replace("淮安", "淮安合计", inplace=True)
    df731["客户"].replace("常州", "常州合计", inplace=True)
    df731["客户"].replace("镇江", "镇江合计", inplace=True)
    df731["客户"].replace("常镇", "常镇合计", inplace=True)
    df731["客户"].replace("绍兴龚群波", "绍兴龚群波合计", inplace=True)
    df731["客户"].replace("衢州", "衢州合计", inplace=True)  #
    df731["客户"].replace("金华", "金华合计", inplace=True)  #
    df731["客户"].replace("无锡1", "无锡1合计", inplace=True)
    df731["客户"].replace("无锡2", "无锡2合计", inplace=True)


    df74 = pd.concat([df72, df73, df731],ignore_index=True)


    df74["客户片区"] = df74["片区"]
    df74["客户片区编码"] = df74["片区编码"]
    df74["编码"] = df74["客户编码"]
    df74["客户名称"] = df74["客户"]
    df74["所属公司"] = df74["公司"]
    df74["应收余额"] = df74["余额"]
    df74["1-30天"] = df74["1-30"]
    df74["31-60天"] = df74["31-60"]
    df74["61-90天"] = df74["61-90"]
    df74["91-120天"] = df74["91-120"]
    df74["121-150天"] = df74["121-150"]
    df74["151-天"] = df74["151-"]
    df74["应收余额"] = df74["余额"]
    #以上是为了列排序

    #df74.to_excel(excel_writer="e:/集团账龄数据.xlsx",
              #     sheet_name="处理",
              #     index=False);

    df75 = df74.drop(['片区','片区编码','客户',"客户编码","公司","余额", "1-30", "31-60", "61-90", "91-120", "121-150","151-"], axis=1)

    df76 = df75.sort_values(by=['序列号码','部门'], axis=0, ascending=True)#行排序


    #df75.sort_valuse(by = ["部门","客户片区"],ascending=[True,False])
    #df74.columns = ['部门','片区','客户',"客户编码","公司","余额", "1-30", "31-60", "61-90", "91-120", "121-150","151-"]
    #order = ['部门','片区','客户',"客户编码","公司","余额", "1-30", "31-60", "61-90", "91-120", "121-150","151-"]
    #dataframe = dataframe[order]
    #df72.loc[片区合计] = df72.groupby(["片区"])["余额", "Unnamed: 8", "Unnamed: 9", "Unnamed: 10", "Unnamed: 11", "Unnamed: 12"].sum();
    #df72.loc['片区合计'] = df72.apply(lambda x: x.sum())
    #df72.loc['01部合计'] = df72[df72["部门"]="01部"].sum()

    print(df76)
   # df76.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630

    df76.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"));
    tkinter.messagebox.showinfo("运行结果", "集团客户账龄分列导出成功！");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())


def appendStr10():  # 药品公司账龄分列&计算账龄
 try:
    tkinter.messagebox.showinfo("提醒", "请先选择金蝶账龄源文件");
    df80 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

    df81 = df80['客户'].str.split('-| ', expand=True);

    df82 = pd.merge(df81, df80, right_index=True, left_index=True);

    df82["客户名称"] = df82[7]
    df82["业务员"] = df82[4]

    df83 = df82.groupby(["客户名称","业务员"], as_index=False)["过期", "Unnamed: 6", "Unnamed: 7", "Unnamed: 8", "Unnamed: 9"].sum();

    c_df = pd.DataFrame(df83)
    c_df.reset_index(inplace=True)  # 取消合并

    df84 = df83.sort_values(by=['业务员'], axis=0, ascending=True)  # 行排序

    df851 = df84.rename(columns={'过期': '1-30', 'Unnamed: 6': '31-60', 'Unnamed: 7': '61-90', 'Unnamed: 8': '91-120',
                         'Unnamed: 9': '121-'});

    tkinter.messagebox.showinfo("提醒", "请选择客户账龄天数源文件");
    #加入账龄表

    df92 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

    df93 = df92.drop(["业务员"], axis=1)
    df94 = df93.rename(columns={'客户': '客户名称'})
    df85 = pd.merge(df851, df94, how='left', on=['客户名称']);  # 完全相同合并，忽略没有的客户(没有how)



    df85["余额"] = df85["1-30"] + df85["31-60"] + df85["61-90"] + df85["91-120"] + df85["121-"]
    df85["序列"] =df85["业务员"]

    df85["序列"].replace("傅诗云", "001", inplace=True)
    df85["序列"].replace("戚肆朝", "002", inplace=True)
    df85["序列"].replace("张丽君", "003", inplace=True)
    df85["序列"].replace("王振杰", "004", inplace=True)
    df85["序列"].replace("杨阳", "005", inplace=True)
    df85["序列"].replace("徐蕾", "006", inplace=True)
    df85["序列"].replace("王伟平", "007", inplace=True)
    df85["序列"].replace("朱津齐", "0082", inplace=True)
    df85["序列"].replace("王宇栋", "0081", inplace=True)
    df85["序列"].replace("张叶", "009", inplace=True)
    df85["序列"].replace("杨利忠", "009", inplace=True)
    df85["序列"].replace("吴优", "009", inplace=True)
    df85["序列"].replace("王建靓", "009", inplace=True)
    df85["序列"].replace("牛宁慧", "009", inplace=True)
    df85["序列"].replace("高波", "009", inplace=True)
    df85["序列"].replace("安铁", "009", inplace=True)
    df85["序列"].replace("冯露敏", "010", inplace=True)
    df85["序列"].replace("余海燕", "010", inplace=True)
    df85["序列"].replace("唐惠", "010", inplace=True)
    df85["序列"].replace("孙婷婷", "010", inplace=True)
    df85["序列"].replace("石亚国", "010", inplace=True)
    df85["序列"].replace("马思远", "010", inplace=True)
    df85["序列"].replace("吕淳昱", "010", inplace=True)
    df85["序列"].replace("金英明", "010", inplace=True)
    df85["序列"].replace("胡迪锋", "010", inplace=True)
    df85["序列"].replace("龚文青", "010", inplace=True)
    df85["序列"].replace("高大勇", "010", inplace=True)
    df85["序列"].replace("丁玲", "010", inplace=True)



    df86 = df85.sort_values(by=['序列'], axis=0, ascending=True)  # 行排序
    print(df86)

    df87 = df86.groupby(["序列"], as_index=False)["余额", "1-30", "31-60", "61-90", "91-120", "121-"].sum();

    df87["客户名称"] = df87["序列"]

    df87["序列"].replace("001", "00199", inplace=True)
    df87["序列"].replace("002", "00299", inplace=True)
    df87["序列"].replace("003", "00399", inplace=True)
    df87["序列"].replace("004", "00499", inplace=True)
    df87["序列"].replace("005", "00599", inplace=True)
    df87["序列"].replace("006", "00699", inplace=True)
    df87["序列"].replace("007", "00799", inplace=True)
    df87["序列"].replace("0082", "008299", inplace=True)
    df87["序列"].replace("0081", "008199", inplace=True)
    df87["序列"].replace("009", "00999", inplace=True)
    df87["序列"].replace("010", "01099", inplace=True)

    df87["客户名称"].replace("001", "傅诗云小计", inplace=True)
    df87["客户名称"].replace("002", "戚肆朝小计", inplace=True)
    df87["客户名称"].replace("003", "张丽君小计", inplace=True)
    df87["客户名称"].replace("004", "王振杰小计", inplace=True)
    df87["客户名称"].replace("005", "杨阳小计", inplace=True)
    df87["客户名称"].replace("006", "徐蕾小计", inplace=True)
    df87["客户名称"].replace("007", "王伟平小计", inplace=True)
    df87["客户名称"].replace("0082", "朱津齐小计", inplace=True)
    df87["客户名称"].replace("0081", "王宇栋小计", inplace=True)
    df87["客户名称"].replace("009", "其他小计", inplace=True)
    df87["客户名称"].replace("010", "试剂小计", inplace=True)
    df87["客户名称"].replace("", "空白小计", inplace=True)

    df88 = pd.concat([df86, df87], ignore_index=True)

    df89 = df88.sort_values(by=['序列'], axis=0, ascending=True)  # 行排序
    print(df89)



    df89["序列号"] = df89["序列"]
    df89["客户"] = df89["客户名称"]
    df89["业务员名称"] = df89["业务员"]
    df89["客户名称"] = df89["客户"]
    df89["应收账龄"] = df89["账龄"]
    df89["应收余额"] = df89["余额"]
    df89["1-30天"] = df89["1-30"]
    df89["31-60天"] = df89["31-60"]
    df89["61-90天"] = df89["61-90"]
    df89["91-120天"] = df89["91-120"]
    df89["121-天"] = df89["121-"]

    # 以上是为了列排序

    df90 = df89.drop(["客户", "1-30","121-","31-60","61-90","91-120","index","余额","序列","业务员","账龄"], axis=1)  # 删列

    df91 = df90.sort_values(by=['序列号'], axis=0, ascending=True)  # 行排序
    print(df91)
     ####以上主表OK,以下是计算逾期账龄704
     ##先分解账龄表拼接到主表

    #df92 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径
   # df93 = df92.drop(["业务员"], axis=1)
   # df94 = df93.rename(columns={'客户': '客户名称'})  #把原来的 客户 命名为 客户名称
       #df94.to_excel(excel_writer="e:/集团账龄数据.xlsx",
        #                 sheet_name="处理",
         #                 index=False);

    #df95 = pd.merge(df91, df94, how='left', on=['客户名称']);  # 完全相同合并，忽略没有的客户(没有how)
    #print(df95)
    #30天的列
    df92=df91[df91["应收账龄"]==30]
    df92["逾期金额"]=df92["1-30天"]+df92["31-60天"]+df92["61-90天"]+df92["91-120天"]+df92["121-天"]
    print(df92)
    df93 = df92.drop(["1-30天", "121-天", "31-60天", "61-90天", "91-120天", "序列号", "业务员名称", "应收账龄","应收余额"], axis=1)  # 删列


    #60天的列
    df94 = df91[df91["应收账龄"] == 60]
    df94["逾期金额"] =df94["31-60天"] + df94["61-90天"] + df94["91-120天"] + df94["121-天"]
    print(df94)
    df95 = df94.drop(["1-30天", "121-天", "31-60天", "61-90天", "91-120天", "序列号", "业务员名称", "应收账龄", "应收余额"], axis=1)  # 删列

    #90天的列
    df96 = df91[df91["应收账龄"] == 90]
    df96["逾期金额"] = df96["61-90天"] + df96["91-120天"] + df96["121-天"]
    print(df96)
    df97 = df96.drop(["1-30天", "121-天", "31-60天", "61-90天", "91-120天", "序列号", "业务员名称", "应收账龄", "应收余额"], axis=1)  # 删列

    #120天的列
    df98 = df91[df91["应收账龄"] == 120]
    df98["逾期金额"] = df98["91-120天"] + df98["121-天"]
    print(df98)
    df99 = df98.drop(["1-30天", "121-天", "31-60天", "61-90天", "91-120天", "序列号", "业务员名称", "应收账龄", "应收余额"], axis=1)  # 删列

    #150天的列
    df100 = df91[df91["应收账龄"] == 150]
    df100["逾期金额"] = df100["121-天"]
    print(df100)
    df101 = df100.drop(["1-30天", "121-天", "31-60天", "61-90天", "91-120天", "序列号", "业务员名称", "应收账龄", "应收余额"], axis=1)  # 删列

    #组合上述列
    df102=pd.concat([df93, df95, df97,df99,df101], ignore_index=True)

    df103 = pd.merge(df91, df102, how='left', on=['客户名称']);  # 完全相同合并，忽略没有的客户(没有how)

    print(df103)
   # df103.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630

    #df103.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
     #                                                               filetypes=[("Microsoft Excel文件", "*.xlsx"),
      #                                                                         (
       #                                                                        "Microsoft Excel 97-20003 文件", "*.xls")],
        #                                                            defaultextension=".xlsx"));

    df103.to_excel("宁波医药客户账龄逾期计算表" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                  index=False)  # 自动输出

    tkinter.messagebox.showinfo("运行结果", "医药客户账龄逾期导出成功！");
 except Exception as error:
    tm.showerror(title="煎饼提示前方路堵",
                 message="请检查提交源文件是否正确 '" + str(error) + "'.",
                 detail=traceback.format_exc())


def appendStr14():#英克货品分类（丘）关联交易
 try:
    tkinter.messagebox.showinfo("提醒","请先选择宁波医药英克源文件");
    df1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径
    df2 = df1.groupby(["货品分类9"], as_index=False)["上期金额","本期金额"].sum();

    #df2["货品分类"]=df2["货品分类9"]
    #df2["货品分类"].replace("", "医药关联货品", inplace=True)

    df3 = df2[df2["货品分类9"] == "非医药类商品"]
    df4 = df2[df2["货品分类9"] == "试剂（器械）"]
    df5 = df2[df2["货品分类9"] == "原厂器械仪器"]
    df6 = df2[df2["货品分类9"] == "进口"]
    df7 = df2[df2["货品分类9"] == "其他"]
    df8 = df2[df2["货品分类9"] == "国产试剂"]
    df9 = df2[df2["货品分类9"] == "国产"]

    df10 = pd.concat([df3, df4, df5, df6, df7, df8, df9],ignore_index=True)
    df10["本期发生"]=df10["本期金额"]-df10["上期金额"]
    df10["货品分类"]='关联交易货品'
    df11 = df10.groupby(["货品分类"], as_index=False)["上期金额", "本期金额","本期发生"].sum();
    df11["所属公司"]='宁波医药'

    print(df11)

    ###医药上传完毕
    ###1.宁波医药 2.浙江3.上海器械4.上海诊断5.江苏恒奇
    tkinter.messagebox.showinfo("提醒", "请选择浙江医疗英克源文件");
    df12 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径
    df13 = df12.groupby(["货品ID"], as_index=False)["上期金额", "本期金额"].sum();

    df13["货品分类"]='关联交易货品'
    df14 = df13.groupby(["货品分类"], as_index=False)["上期金额", "本期金额"].sum();
    df14["本期发生"] = df14["本期金额"] - df14["上期金额"]
    df14["所属公司"] = '浙江医疗'
    print(df14)

    ####浙江上传完毕
    tkinter.messagebox.showinfo("提醒", "请选择上海器械英克源文件");
    df15 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径
    df16 = df15.groupby(["货品ID"], as_index=False)["上期金额", "本期金额"].sum();

    df16["货品分类"] = '关联交易货品'
    df17 = df16.groupby(["货品分类"], as_index=False)["上期金额", "本期金额"].sum();
    df17["本期发生"] = df17["本期金额"] - df17["上期金额"]
    df17["所属公司"] = '上海器械'
    print(df17)

    ####上海器械完毕
    tkinter.messagebox.showinfo("提醒", "请选择上海诊断英克源文件");
    df18 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径
    df19 = df18.groupby(["货品ID"], as_index=False)["上期金额", "本期金额"].sum();

    df19["货品分类"] = '关联交易货品'
    df20 = df19.groupby(["货品分类"], as_index=False)["上期金额", "本期金额"].sum();
    df20["本期发生"] = df20["本期金额"] - df20["上期金额"]
    df20["所属公司"] = '上海诊断'
    print(df20)

    ####上海诊断完毕
    tkinter.messagebox.showinfo("提醒", "请选择江苏恒奇诊断英克源文件");
    df21 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径
    df22 = df21.groupby(["货品ID"], as_index=False)["上期金额", "本期金额"].sum();

    df22["货品分类"] = '关联交易货品'
    df23 = df22.groupby(["货品分类"], as_index=False)["上期金额", "本期金额"].sum();
    df23["本期发生"] = df23["本期金额"] - df23["上期金额"]
    df23["所属公司"] = '江苏恒奇诊断'
    print(df23)
    ####江苏恒奇诊断完毕


    #####合并

    df100 = pd.concat([df11,df14,df17,df20,df23],ignore_index=True)
    df101 = df100.groupby(["货品分类"], as_index=False)["上期金额", "本期金额","本期发生"].sum();
    df102 = pd.concat([df100, df101], ignore_index=True)
    df102["所属公司1"] = df102["所属公司"]
    df102["上期金额1"]=df102["上期金额"]
    df102["本期金额1"] = df102["本期金额"]
    df102["本期变动额"] = df102["本期发生"]

    df103 = df102.drop(["本期金额","所属公司","上期金额","本期发生"], axis=1)  # 删列
    df104 = df103.rename(columns={'本期金额1': '本期金额','所属公司1':'所属公司','上期金额1':'上期金额'})
    df105 = df104.fillna('合计')
    #df104["所属公司"].replace(" ", "合计", inplace=True)


    print(df105)

    #df3 = df2.groupby(["货品分类"], as_index=False)["上期金额", "本期金额"].sum();

    #tkinter.messagebox.showinfo("1");
    #df105.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630

    df105.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"));
    tkinter.messagebox.showinfo("运行结果", "货品关联交易查询表导出成功！");

########英克货品关联抵消表完毕
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())



def appendStr16():  #对账单自动分客户
 try:
    tkinter.messagebox.showinfo("提醒", "请先选择金蝶对账源文件");
    data1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径

    c_df = pd.DataFrame(data1);
    c_df.reset_index(inplace=True);

    df1=data1.drop(data1.index[[0,1]], axis=0)
    df2=df1['科目：1122 - 应收账款'].str.split('-| ', expand=True);

    df3 =pd.merge(df2, df1, right_index=True, left_index=True);


    df4=df3.rename(columns={0: '部门',1: '片区',7: '客户','币别:人民币':'项目','Unnamed: 3':'日期','Unnamed: 5':'摘要','Unnamed: 6':'发票号','Unnamed: 8':'借方','Unnamed: 9':'贷方','Unnamed: 10':'方向','Unnamed: 11':'余额'});

    #df3["部门"]=df3[0]
    #df3["片区"]=df3[1]
    #df3["客户"]=df3[7]
    #df3["摘要1"]=df3[14]
    #df3["借方1"]=df3[17]
    df5=df4.drop([2, 3, 4, 5,6,"科目：1122 - 应收账款","Unnamed: 1"], axis=1)  # 删列

    # df51=df5.iloc[:,[7,8]].astype("float64")
    df51=df5.fillna(0)
    df52=df51[df51["客户"] != 0]
    df53=df52.rename(columns={'发票号': '借', '借方': '方向', '贷方': '余额',5:'贷'});
    # df54=df53["借"].astype("float64")
    # df52["客户"].replace("8", "无", inplace=True)

    #df4=df4.drop(df4.columns[[[[[[[[0,1,2,3,4,5,6,7,8,9,10]]]]]]]], axis=1)  # 删列

    #df5 = df4.fillna(0)
    #df6 = df5[df5["客户"] != 0]
    # df7.to_excel(excel_writer="D:/自动回复测试/对账客户初步分离.xlsx",
    #                   sheet_name="处理",
    #                  index=False);

    #df6["项目"]=df6["Unnamed: 1"]
   # df6["日期"]=df6["币别:人民币"]
    #df6["凭证号"] = df6["Unnamed: 3"]
   # #df6["借方"] = df6["Unnamed: 6"]
    #df6["贷方"] = df6["Unnamed: 8"]
   # df6["期末余额"] = df6["Unnamed: 10"]
   # df6["期末方向"] = df6["Unnamed: 9"]


   # df7 =df6.drop(["Unnamed: 8","Unnamed: 1","Unnamed: 3","Unnamed: 9","Unnamed: 6","Unnamed: 5","币别:人民币","Unnamed: 10"], axis=1)  # 删列

    #df7.to_excel(excel_writer="D:/自动回复测试/对账客户初步分离.xlsx",
     #                   sheet_name="处理",
      #                  index=False);

    df53.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630

    tkinter.messagebox.showinfo("提醒", "请先选择刚才输出的文件");

    data=pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径

    rows = data.shape[0]  # 获取行数 shape[1]获取列数
    department_list = []
    for i in range(rows):
        temp = data["客户"][i]
        if temp not in department_list:
            department_list.append(temp)  # 将销售部门的分类存在一个列表中

    for department in department_list:
        new_df = pd.DataFrame()

        for i in range(0, rows):
            if data["客户"][i] == department:
                new_df = pd.concat([new_df, data.iloc[[i], :]], axis=0, ignore_index=True)

        new_df.to_excel(str(department) + ".xls", sheet_name=department, index=False)  # 将每个销售部门存成一个新excel

    tkinter.messagebox.showinfo("运行结果", "客户对账单整理成功！");
    #df3.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())

def appendStr17():  #货品成本分析
 try:
    tkinter.messagebox.showinfo("提醒", "请先选择生物英克进销存源文件");
    df23 = pd.read_excel(tkinter.filedialog.askopenfilename());

    df24 = df23.groupby(["货品ID","货品分类8","通用名"])["本期金额", "本期数量", "进货金额", "进货数量", "上期金额", "上期数量", "调整金额", "销售发票数量", "销售发票金额"].sum();

    df24['成本单价(生物)']= (df24['本期金额'] + df24['上期金额'] + df24['进货金额'] + df24["销售发票金额"] - df24['调整金额']) / (
                df24['本期数量'] + df24['上期数量'] + df24['进货数量'] + df24["销售发票数量"])  # 后面增加一列相除单价



    print(df24);


    tkinter.messagebox.showinfo("提醒", "请先选择英克销售发票明细源文件");
    df25 = pd.read_excel(tkinter.filedialog.askopenfilename());

    df25["序列"] = df25["货品ID"]
    df26 = df25.groupby(["货品ID","通用名"])["基本单位数量","销售成本","价额"].sum();  #要修改为销售发票明细8.8



    df27 = pd.merge(df26, df24, how='left', on=['货品ID','通用名']);  # 完全相同合并，忽略没有的货品ID(没有how)

    c_df = pd.DataFrame(df27)
    c_df.reset_index(inplace=True)  # 取消合并

    df271 = df27[df27["价额"] != 0]

    df271['成本(生物)'] = df271['成本单价(生物)'] * df271['基本单位数量']
    df271['毛利率'] = (df271['价额'] - df271['销售成本']) / df271['价额']
    df271['毛利率(生物)'] = (df271['价额']-df271['成本(生物)'])/df271['价额']

    df271['成本(生物)'].astype("float64")
    df271['毛利率'].astype("float64")
    df271['毛利率(生物)'].astype("float64")


    df30 = df27[df27["价额"] == 0]

    df30['成本(生物)'] = df30['成本单价(生物)'] * df30['基本单位数量']
    df30['毛利率'] = 0
    df30['毛利率(生物)'] = 0

    df30['成本(生物)'].astype("float64")
    df30['毛利率'].astype("float64")
    df30['毛利率(生物)'].astype("float64")



    df31 = pd.concat([df271, df30], ignore_index=True)



    df32 = df31.fillna(0)
    #c_df = pd.DataFrame(df30);
   # c_df.reset_index(inplace=True)


    df32["毛利"]=df32["价额"]-df32["销售成本"]
    df32["毛利(生物)"]=df32["价额"]-df32["成本(生物)"]


    df33 = df32.drop(["本期金额","本期数量","进货金额","进货数量","上期金额","上期数量","调整金额","销售发票数量","销售发票金额"],axis = 1)#删列

    print(df33)

    df34 = df33.fillna(0)
    #df33.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='*.xlsx', ));  # 指定位置另存为630

    df34.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                           ("Microsoft Excel 97-20003 文件", "*.xls")],
                                                                defaultextension=".xlsx"));


    tkinter.messagebox.showinfo("运行结果", "货品成本分析导出成功！");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())



def appendStr18():  # 客户成本分析
 try:
    tkinter.messagebox.showinfo("提醒", "请先选择生物英克进销存源文件");

    df23 = pd.read_excel(tkinter.filedialog.askopenfilename());


    df24 = df23.groupby(["货品ID", "货品分类8", "通用名"])[
        "本期金额", "本期数量", "进货金额", "进货数量", "上期金额", "上期数量", "调整金额", "销售发票数量", "销售发票金额"].sum();

    df24['成本单价(生物)'] = (df24['本期金额'] + df24['上期金额'] + df24['进货金额'] + df24["销售发票金额"] - df24['调整金额']) / (
            df24['本期数量'] + df24['上期数量'] + df24['进货数量'] + df24["销售发票数量"])  # 后面增加一列相除单价

    print(df24);

    tkinter.messagebox.showinfo("提醒", "请先选择英克销售发票明细源文件");
    df25 = pd.read_excel(tkinter.filedialog.askopenfilename());

    df25["序列"] = df25["货品ID"]
    df26 = df25.groupby(["客户","货品ID", "通用名"])["基本单位数量", "销售成本", "价额"].sum();  # 要修改为销售发票明细8.8

    c_df = pd.DataFrame(df26)
    c_df.reset_index(inplace=True)  # 取消合并

    df27 = pd.merge(df26, df24, how='left', on=['货品ID', '通用名']);  # 完全相同合并，忽略没有的货品ID(没有how)

    c_df = pd.DataFrame(df27)
    c_df.reset_index(inplace=True)  # 取消合并

   # df27.to_excel(excel_writer="D:/自动回复测试/对账客户初步分离.xlsx",
    #                                sheet_name="处理",
     #                               index=False);



    df28 = df27.drop(["本期金额","本期数量","进货金额","进货数量","上期金额","上期数量","调整金额","销售发票数量","销售发票金额","index"],axis = 1)#删列

    df28["成本(生物)"]=df28["基本单位数量"]*df28["成本单价(生物)"]
    df28['毛利率'] = (df28['价额'] - df28['销售成本']) / df28['价额']
    df28['毛利率(生物)'] = (df28['价额'] - df28['成本(生物)']) / df28['价额']

    df28["毛利"] = df28["价额"] - df28["销售成本"]
    df28["毛利(生物)"] = df28["价额"] - df28["成本(生物)"]

    df28['成本(生物)'].astype("float64")
    df28['毛利率'].astype("float64")
    df28['毛利率(生物)'].astype("float64")

    df29 = df28.fillna(0)
    df30 = df29.groupby(["客户"])[
        "价额", "销售成本", "成本(生物)", "毛利", "毛利(生物)"].sum();

    c_df = pd.DataFrame(df30)
    c_df.reset_index(inplace=True)  # 取消合并



    print(df29)
    print(df30)



    #df29.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb' ,defaultextension='*.xlsx', ));  # 指定位置另存为630

    df30.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                       filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                  ("Microsoft Excel 97-20003 文件", "*.xls")],
                                       defaultextension=".xlsx"));
    df29.to_excel("客户成本带货品明细" +str(datetime.datetime.now().strftime('%Y%m%d'))+ ".xls", sheet_name="sheet1", index=False) #自动输出
    tkinter.messagebox.showinfo("运行结果", "客户成本分析导出成功！");

 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())

def appendStr19():  #######客户账龄分级婷婷

 try:
    tkinter.messagebox.showinfo("提醒", "请先选择账龄源文件");
    df23 = pd.read_excel(tkinter.filedialog.askopenfilename());



    df24 = df23['客户'].str.split('-| ', expand=True);


    #df241 = df24[df24["公司"] != "海尔施集团"]


    df25 = pd.merge(df24, df23, right_index=True, left_index=True);

    df25["客户名称"] = df25[7]
    df261 = df25[df25["公司"] != "海尔施集团"]
    df262 = df261[df261["项目"] == "001 代理试剂"]

    df26 = df262.groupby(["客户名称"], as_index=False)[
        "过期", "Unnamed: 8", "Unnamed: 9", "Unnamed: 10", "Unnamed: 11","Unnamed: 12","Unnamed: 13","Unnamed: 14","Unnamed: 15","Unnamed: 16"].sum();


    c_df = pd.DataFrame(df26)
    c_df.reset_index(inplace=True)  # 取消合并

    df27 = df26.rename(columns={'过期': '1-7', 'Unnamed: 8': '8-30', 'Unnamed: 9': '31-60', 'Unnamed: 10': '61-90','Unnamed: 11': '91-105',
                                 'Unnamed: 12': '106-120','Unnamed: 13': '121-150','Unnamed: 14': '151-180','Unnamed: 15': '181-300','Unnamed: 16': '301-'});



    ######上面处理账龄，下面开始合并协议天数
    tkinter.messagebox.showinfo("提醒", "请选择客户协议天数源文件");
    # 加入协议天数

    df92 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径


    df93 = pd.merge(df92, df27, how='left', on=['客户名称']);  # 完全相同合并，忽略没有的客户(没有how)

    df94 = df93.drop(["客户性质"], axis=1)




    df95 = df94.drop(["index"], axis=1)



    #7天的列
    df971 = df95[df95["天数"] == 7]
    df971["逾期金额"] = df971["8-30"] +df971["31-60"] + df971["61-90"] + df971["91-105"] + df971["106-120"] + df971["121-150"] + df971[
        "151-180"] + df971["181-300"] + df971["301-"]
    print(df971)
    df981 = df971.drop(["1-7","8-30", "31-60", "61-90", "91-105", "106-120", "121-150", "151-180", "181-300", "301-"],
                     axis=1)  # 删列



    # 30天的列
    df97 = df95[df95["天数"] == 30]
    df97["逾期金额"] = df97["31-60"] + df97["61-90"] + df97["91-105"] + df97["106-120"] + df97["121-150"]+ df97["151-180"]+ df97["181-300"]+ df97["301-"]
    print(df97)
    df98 = df97.drop(["1-7","8-30", "31-60", "61-90", "91-105", "106-120", "121-150", "151-180", "181-300", "301-"], axis=1)  # 删列

    # 60天的列
    df99 = df95[df95["天数"] == 60]
    df99["逾期金额"] =df99["61-90"] + df99["91-105"] + df99["106-120"] + df99["121-150"] + df99[
        "151-180"] + df99["181-300"] + df99["301-"]
    print(df99)
    df100 = df99.drop(["1-7","8-30", "31-60", "61-90", "91-105", "106-120", "121-150", "151-180", "181-300", "301-"],
                     axis=1)  # 删列

    # 90天的列
    df101 = df95[df95["天数"] == 90]

    df101["逾期金额"] = df101["91-105"] + df101["106-120"] + df101["121-150"] + df101[
        "151-180"] + df101["181-300"] + df101["301-"]
    print(df101)
    df102 = df101.drop(["1-7","8-30", "31-60", "61-90", "91-105", "106-120", "121-150", "151-180", "181-300", "301-"],
                      axis=1)  # 删列

    # 105天的列
    df103 = df95[df95["天数"] == 105]

    df103["逾期金额"] = df103["106-120"] + df103["121-150"] + df103["151-180"] + df103["181-300"] + df103["301-"]
    print(df103)
    df104 = df103.drop(["1-7","8-30", "31-60", "61-90", "91-105", "106-120", "121-150", "151-180", "181-300", "301-"],
                       axis=1)  # 删列

    # 120天的列
    df106 = df95[df95["天数"] == 120]

    df106["逾期金额"] =df106["121-150"] + df106["151-180"] + df106["181-300"] + df106["301-"]
    print(df106)
    df107 = df106.drop(["1-7","8-30", "31-60", "61-90", "91-105", "106-120", "121-150", "151-180", "181-300", "301-"],
                       axis=1)  # 删列

    # 150天的列
    df109 = df95[df95["天数"] == 150]

    df109["逾期金额"] = df109["151-180"] + df109["181-300"] + df109["301-"]
    print(df109)
    df110 = df109.drop(["1-7","8-30", "31-60", "61-90", "91-105", "106-120", "121-150", "151-180", "181-300", "301-"],
                       axis=1)  # 删列

    # 180天的列
    df112 = df95[df95["天数"] == 180]

    df112["逾期金额"] = df112["181-300"] + df112["301-"]
    print(df112)
    df113 = df112.drop(["1-7","8-30", "31-60", "61-90", "91-105", "106-120", "121-150", "151-180", "181-300", "301-"],
                       axis=1)  # 删列

    # 300天的列
    df115 = df95[df95["天数"] == 300]

    df115["逾期金额"] = df115["301-"]
    print(df115)
    df116 = df115.drop(["1-7","8-30", "31-60", "61-90", "91-105", "106-120", "121-150", "151-180", "181-300", "301-"],
                       axis=1)  # 删列

    # 20万货款的列
    df118 = df95[df95["天数"] == '20w货款账期']

    df118["逾期金额"] = df118["1-7"]+df118["8-30"]+df118["31-60"] + df118["61-90"] + df118["91-105"] + df118["106-120"] + df118["121-150"]+ df118["151-180"]+ df118["181-300"]+ df118["301-"]-200000
    print(df118)
    df120 = df118.drop(["1-7","8-30", "31-60", "61-90", "91-105", "106-120", "121-150", "151-180", "181-300", "301-"],
                       axis=1)  # 删列

    # 0天的列
    df121 = df95[df95["天数"] == 0]

    df121["逾期金额"] = df121["1-7"]+df121["8-30"] + df121["31-60"] + df121["61-90"] + df121["91-105"] + df121["106-120"] + df121[
        "121-150"] + df121["151-180"] + df121["181-300"] + df121["301-"]
    print(df121)
    df122 = df121.drop(["1-7","8-30", "31-60", "61-90", "91-105", "106-120", "121-150", "151-180", "181-300", "301-"],
                       axis=1)  # 删列




     ###拼接逾期金额

    df121 = pd.concat([df981,df98, df100, df102, df104, df107,df110,df113,df116,df120,df122], ignore_index=True)

    df122 =df95.drop(["部门", "区域", "天数"],
                       axis=1)

    df123 = pd.merge(df121, df122, how='left', on=['客户名称']);  # 完全相同合并，忽略没有的客户(没有how)

    df124 = df123[df123["逾期金额"] > 0]

    tkinter.messagebox.showinfo("提醒", "请选择金蝶科目项目余额表源文件");
    # 加入应收账款借方本年发生额

    df150 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

    c_df = pd.DataFrame(df150)
    c_df.reset_index(inplace=True)

    df151= df150.drop([0],axis=0)
    df152 = df151[df151["项目"] == "代理试剂"]
    df153 = df152[df152["公司"] != "海尔施集团"]

    df154 = df153.groupby(["客户"], as_index=False)["本年累计"].sum();

    df155 = df154.rename(columns={'客户': '客户名称'});

    df156 = pd.merge(df124, df155, how='left', on=['客户名称']);  # 完全相同合并，忽略没有的客户(没有how)

    #####自动生成完全版数据

    #df157 = pd.merge(df123, df155, how='left', on=['客户名称']);

    #df151 = df150['项目代码'].str.split('-| ', expand=True);

    #df152 = pd.merge(df151, df23, right_index=True, left_index=True);

    df261 = df25[df25["公司"] != "海尔施集团"]
    df262 = df261[df261["项目"] == "001 代理试剂"]

    ####19-10-25整理输出表格格式

    #df157 = df156.drop(["Unnamed: 6_x", "等级_y", "Unnamed: 6_y"], axis=1)

    #df158 = df157.rename(columns={'等级_x': '等级'});



    df156.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               ("Microsoft Excel 97-20003 文件",
                                                                                   "*.xls")],
                                                                    defaultextension=".xlsx"));

   # df262.to_excel("客户等级全部公司版" + str(datetime.datetime.now().strftime('%Y%m%d%h')) + ".xls", sheet_name="sheet1",index=False)  # 自动输出
    tkinter.messagebox.showinfo("运行结果", "客户等级测试导出成功！");

 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
             message="请检查提交源文件是否正确 '" + str(error) + "'.",
             detail=traceback.format_exc())




def appendStr20():  ##########################################################################################销售发票分税率巧儿

 try:
    tkinter.messagebox.showinfo("提醒", "请先选择开票明细源文件");
    df = pd.read_excel(tkinter.filedialog.askopenfilename());

    df1 = df.drop(df.columns[[[[[[[[[[3, 4, 5, 7, 8, 10, 11, 12, 13, 17]]]]]]]]]], axis=1);
    df2 = df1.drop(df1.index[[[[0, 1, 2, 3]]]], axis=0);
    df3 = df2[df2["Unnamed: 9"] != "小计"];
    df4 = df3[df3["Unnamed: 9"] != "商品名称"];
    df5 = df4.dropna(how="all");
    df6 = df5.fillna(method='pad');
    df6["Unnamed: 14"] = df6["Unnamed: 14"].astype("float64");  # 改变格式
    df6["Unnamed: 16"] = df6["Unnamed: 16"].astype("float64");

    df9 = df6.loc[df6['Unnamed: 9'].str.contains('紫杉醇注射液')]  # 7-15模糊查询,但单元格不能为0为空  （ #print(df4[df4['客户'].isin(['宁波海尔施医学检验所有限公司'])])）
    df10 = df6.loc[df6['Unnamed: 9'].str.contains('阿那曲唑片')]
    df11 = df6.loc[df6['Unnamed: 9'].str.contains('奥沙利铂甘露醇注射液')]
    df12 = df6.loc[df6['Unnamed: 9'].str.contains('比卡鲁胺片')]
    df13 = df6.loc[df6['Unnamed: 9'].str.contains('醋酸奥曲肽注射液')]
    df14 = df6.loc[df6['Unnamed: 9'].str.contains('醋酸戈舍瑞林缓释植入剂')]
    df15 = df6.loc[df6['Unnamed: 9'].str.contains('多西他赛注射液')]
    df16 = df6.loc[df6['Unnamed: 9'].str.contains('吉非替尼片')]
    df17 = df6.loc[df6['Unnamed: 9'].str.contains('甲苯磺酸索拉非尼片')]
    df18 = df6.loc[df6['Unnamed: 9'].str.contains('甲磺酸奥希替尼片')]
    df19 = df6.loc[df6['Unnamed: 9'].str.contains('甲磺酸伊马替尼片')]
    df20 = df6.loc[df6['Unnamed: 9'].str.contains('酒石酸长春瑞滨注射液')]
    df21 = df6.loc[df6['Unnamed: 9'].str.contains('卡培他滨片')]
    df22 = df6.loc[df6['Unnamed: 9'].str.contains('来曲唑片')]
    df23 = df6.loc[df6['Unnamed: 9'].str.contains('硫培非格司亭注射液')]
    df24 = df6.loc[df6['Unnamed: 9'].str.contains('顺铂注射液')]
    df25 = df6.loc[df6['Unnamed: 9'].str.contains('替吉奥胶囊')]
    df26 = df6.loc[df6['Unnamed: 9'].str.contains('注射用奥沙利铂')]
    df27 = df6.loc[df6['Unnamed: 9'].str.contains('注射用地西他滨')]
    df28 = df6.loc[df6['Unnamed: 9'].str.contains('注射用洛铂')]
    df29 = df6.loc[df6['Unnamed: 9'].str.contains('注射用奈达铂')]
    df30 = df6.loc[df6['Unnamed: 9'].str.contains('注射用培美曲塞二钠')]
    df31 = df6.loc[df6['Unnamed: 9'].str.contains('注射用亚叶酸钙')]
    df32 = df6.loc[df6['Unnamed: 9'].str.contains('注射用盐酸表柔比星')]
    df33 = df6.loc[df6['Unnamed: 9'].str.contains('注射用盐酸吉西他滨')]
    df34 = df6.loc[df6['Unnamed: 9'].str.contains('注射用盐酸伊立替康')]
    df35 = df6.loc[df6['Unnamed: 9'].str.contains('注射用紫杉醇')]
    df36 = df6.loc[df6['Unnamed: 9'].str.contains('比卡鲁胺胶囊')]

    df40 = pd.concat([df9, df10, df11, df12, df13, df14, df15, df16, df17, df18, df19, df20, df21, df22,df23,
                      df24,df25,df26,df27,df28,df29,df30,df31,df32,df33,df34,df35,df36],
                      ignore_index=True)  # 组合
    df41 = df40.sort_values(by=['Unnamed: 1'], axis=0, ascending=True)  # 行排序

    df42 = df41.rename(columns={'Unnamed: 1': '发票号码', 'Unnamed: 2': '客户', 'Unnamed: 6': '发票日期', 'Unnamed: 9': '货品名称',
                         'Unnamed: 14': '无税金额', 'Unnamed: 15': '税率','Unnamed: 16': '税额'});


    df42["分类"]=df42["税率"]
    df42["分类"].replace("3%", "抗癌3%", inplace=True)




    #######上面是抗癌药物3%
    #####下面开始药品3%
    df49 = df6[df6["Unnamed: 15"] == "3%"];

    df50 = df49.loc[df49['Unnamed: 9'].str.contains('重组人干扰素a2b注射液')]  # 7-15模糊查询,但单元格不能为0为空
    df51 = df49.loc[df49['Unnamed: 9'].str.contains('注射用鼠神经生长因子')]  # 7-15模糊查询,但单元格不能为0为空
    df52 = df49.loc[df49['Unnamed: 9'].str.contains('脑苷肌肽注射液')]  # 7-15模糊查询,但单元格不能为0为空
    df53 = df49.loc[df49['Unnamed: 9'].str.contains('骨瓜提取物注射液')]  # 7-15模糊查询,但单元格不能为0为空
    df54 = df49.loc[df49['Unnamed: 9'].str.contains('注射用骨肽')]  # 7-15模糊查询,但单元格不能为0为空
    df55 = df49.loc[df49['Unnamed: 9'].str.contains('静注人免疫球蛋白')]  # 7-15模糊查询,但单元格不能为0为空
    df56 = df49.loc[df49['Unnamed: 9'].str.contains('人血白蛋白')]  # 7-15模糊查询,但单元格不能为0为空
    df57 = df49.loc[df49['Unnamed: 9'].str.contains('人凝血酶原复合物')]  # 7-15模糊查询,但单元格不能为0为空
    df58 = df49.loc[df49['Unnamed: 9'].str.contains('破伤风人免疫球蛋白')]  # 7-15模糊查询,但单元格不能为0为空
    df59 = df49.loc[df49['Unnamed: 9'].str.contains('缩宫素注射液')]  # 7-15模糊查询,但单元格不能为0为空
    df60 = df49.loc[df49['Unnamed: 9'].str.contains('注射用硼替佐米')]  # 7-15模糊查询,但单元格不能为0为空
    df61 = df49.loc[df49['Unnamed: 9'].str.contains('注射用白眉蛇毒血凝酶')]  # 7-15模糊查询,但单元格不能为0为空
    df62 = df49.loc[df49['Unnamed: 9'].str.contains('酪酸梭菌活菌胶囊')]  # 7-15模糊查询,但单元格不能为0为空

    df63 = pd.concat([df50,df51,df52,df53,df54,df55,df56,df57,df58,df59,df60,df61,df62],ignore_index=True)  # 组合
    print(df63)

    df64 = df63.sort_values(by=['Unnamed: 1'], axis=0, ascending=True)  # 行排序

    df65 = df64.rename(columns={'Unnamed: 1': '发票号码', 'Unnamed: 2': '客户', 'Unnamed: 6': '发票日期', 'Unnamed: 9': '货品名称',
                                'Unnamed: 14': '无税金额', 'Unnamed: 15': '税率', 'Unnamed: 16': '税额'});

    df65["分类"] = df65["税率"]
    df65["分类"].replace("3%", "药品3%", inplace=True)



    #####药品3%结束
    #####试剂3%开始
    df66 = df6[df6["Unnamed: 15"] == "3%"];

    df67 = df66.loc[df66['Unnamed: 9'].str.contains('A抗A抗B血型定型试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df68 = df66.loc[df66['Unnamed: 9'].str.contains('B抗A抗B血型定型试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df69 = df66.loc[df66['Unnamed: 9'].str.contains('抗A,抗B血型定型试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df70 = df66.loc[df66['Unnamed: 9'].str.contains('0605005人类免疫缺陷病毒抗体诊断试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df71 = df66.loc[df66['Unnamed: 9'].str.contains('0605007乙型肝炎病毒表面抗原诊断试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df72 = df66.loc[df66['Unnamed: 9'].str.contains('A 抗A抗B血型定型试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df73 = df66.loc[df66['Unnamed: 9'].str.contains('B 抗A抗B血型定型试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df74 = df66.loc[df66['Unnamed: 9'].str.contains('抗人球蛋白检测卡')]  # 7-15模糊查询,但单元格不能为0为空
    df75 = df66.loc[df66['Unnamed: 9'].str.contains('梅毒螺旋体抗体诊断试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df76 = df66.loc[df66['Unnamed: 9'].str.contains('19211人类免疫缺陷病毒抗体诊断试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df77 = df66.loc[df66['Unnamed: 9'].str.contains('乙型肝炎病毒核心抗体IgM检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df78 = df66.loc[df66['Unnamed: 9'].str.contains('乙型肝炎病毒前S1抗原检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df79 = df66.loc[df66['Unnamed: 9'].str.contains('丙型肝炎病毒抗体诊断试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df80 = df66.loc[df66['Unnamed: 9'].str.contains('梅毒甲苯胺红不加热血清试验诊断试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df81 = df66.loc[df66['Unnamed: 9'].str.contains('ABO血型反定型试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df82 = df66.loc[df66['Unnamed: 9'].str.contains('ABO、RhD血型定型检测卡')]  # 7-15模糊查询,但单元格不能为0为空
    df83 = df66.loc[df66['Unnamed: 9'].str.contains('甲型肝炎病毒IgM抗体检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df84 = df66.loc[df66['Unnamed: 9'].str.contains('A乙型肝炎病毒表面抗体检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df85 = df66.loc[df66['Unnamed: 9'].str.contains('01200205A乙型肝炎病毒e抗原检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df86 = df66.loc[df66['Unnamed: 9'].str.contains('01200208A乙型肝炎病毒核心抗体检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df87 = df66.loc[df66['Unnamed: 9'].str.contains('01200210A乙型肝炎病毒e抗体检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空

    df95 = pd.concat([df67,df68,df69,df70,df71,df72,df73,df74,df75,df76,df77,df78,df79,df80,df81,df82,df83,df84,df85,df86,df87],
                     ignore_index=True)  # 组合
    print(df95)

    df96 = df95.sort_values(by=['Unnamed: 1'], axis=0, ascending=True)  # 行排序

    df97 = df96.rename(columns={'Unnamed: 1': '发票号码', 'Unnamed: 2': '客户', 'Unnamed: 6': '发票日期', 'Unnamed: 9': '货品名称',
                                'Unnamed: 14': '无税金额', 'Unnamed: 15': '税率', 'Unnamed: 16': '税额'});

    df97["分类"] = df97["税率"]
    df97["分类"].replace("3%", "试剂3%", inplace=True)

   ####试剂3%完毕
   ####其他3%

    df110 = df6[df6["Unnamed: 15"] == "3%"];

    A=df110['Unnamed: 9'].str.contains('A抗A抗B血型定型试剂')

    print(A)
    df111 = df110[df110["Unnamed: 9"] != A]




    df112= df111.sort_values(by=['Unnamed: 1'], axis=0, ascending=True)  # 行排序

    df113 = df112.rename(columns={'Unnamed: 1': '发票号码', 'Unnamed: 2': '客户', 'Unnamed: 6': '发票日期', 'Unnamed: 9': '货品名称',
                                'Unnamed: 14': '无税金额', 'Unnamed: 15': '税率', 'Unnamed: 16': '税额'});

    df100 =pd.concat([df42,df65,df97,df113],
                     ignore_index=True)  # 组合

    df7a= df6.rename(columns={'Unnamed: 1': '发票号码', 'Unnamed: 2': '客户', 'Unnamed: 6': '发票日期', 'Unnamed: 9': '货品名称',
                                'Unnamed: 14': '无税金额', 'Unnamed: 15': '税率', 'Unnamed: 16': '税额'});
    df102 = pd.concat([df100,df7a],ignore_index=True)  # 组合
    # 查看是否有重复行
    re_row = df100.duplicated()
    print(re_row)

    # 查看去除重复行的数据
    no_re_row = df100.drop_duplicates()
    print(no_re_row)

    # 查看基于[物品]列去除重复行的数据f
    df101 = df102.drop_duplicates(['货品名称','无税金额','客户','发票号码'])
    #print(wp)
    df101["金额"]=df101["无税金额"]+df101["税额"]
    df101.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                                   "Microsoft Excel 97-20003 文件",
                                                                                   "*.xls")],
                                                                    defaultextension=".xlsx"));


    #df101.to_excel(excel_writer="d:/英克测试.xlsx",
     #             sheet_name="测试1",
      #            );
    tkinter.messagebox.showinfo("运行结果", "开票税率分类导出成功！");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
             message="请检查提交源文件是否正确 '" + str(error) + "'.",
             detail=traceback.format_exc())





def appendStr21():  #######英克销售和开票对比

 try:
    tkinter.messagebox.showinfo("提醒", "请先选择出库明细源文件");
    df1 = pd.read_excel(tkinter.filedialog.askopenfilename());
    df2=df1.fillna(0)
    df3 = df2.drop(df2.index[[0, 1]], axis=0);

    df3["仪器出库合计"]=df3["03原厂仪器"]+df3["04采购平台仪器"]+df3["0501国产辅助配置"]+df3["0502流水线辅助配置"]

    df3["非仪器出库合计"]=df3["010101免疫（代理）"]+df3["010102特定蛋白（代理）"]+df3["010103血球（代理）"]+df3["010104普通生化（代理）"]\
    +df3["010105AU生化（代理）"]+df3["010106利德曼生化（代理）"]+df3["010107尿液（代理）"]+df3["010109微生物（代理）"]+df3["010110索灵（代理BC）"]\
    +df3["010111免疫（AMH）"]+df3["010201血凝（代理）"]+df3["0103lmmucor"]+df3["0104索灵"]+df3["010501质控试剂（代理）"]+df3["010502伯乐其它试剂"]\
    +df3["010601BNP试剂（代理）"]+df3["010701血气（代理）"]+df3["0108苏医（代理BC血球质控）"]+df3["020101干式生化"]+df3["020102普通生化"]\
    +df3["020103血气"]+df3["020104特殊生化"]+df3["020201血球"]+df3["020202血凝"]+df3["020203尿液"]+df3["020204血库"]+df3["020206体液"]\
    +df3["020301发光"] +df3["020302特定蛋白"]+df3["020303酶免类"]+df3["020304其它免疫"]+df3["020305厦门万泰"]+df3["0204微生物"]+df3["0205药字号"]\
    +df3["0206分子诊断"]+df3["0207病理科"]+df3["0208采购平台其它"]+df3["0209质控"]+df3["06软件"]+df3["07配件"] \
    +df3["08其它业务"]+df3["0901基因试剂（自产）"]+df3["0902基因试剂（其它厂家）"]+df3["1101强盛生化"]+df3["1201沃文特免疫"]+df3["1202沃文特其他"]\
    +df3["99其它"]

    df4 = df3.drop(["Unnamed: 1", "Unnamed: 2", "Unnamed: 3"], axis=1)  # 删列

    df5 = df4.groupby(["Unnamed: 0"], as_index=False)["非仪器出库合计", "仪器出库合计"].sum();

    #df5["Unnamed: 0"]=df5[""]
    df5["地区"] = df5["Unnamed: 0"]
    df5["负责人"] = df5["Unnamed: 0"]
    df5["地区编码"] = df5["Unnamed: 0"]
    df5["部门"] = df5["Unnamed: 0"]


    df5["地区"].replace("温州葛瑞", "温州1", inplace=True)
    df5["地区编码"].replace("温州葛瑞", "0101", inplace=True)
    df5["负责人"].replace("温州葛瑞", "葛瑞", inplace=True)
    df5["部门"].replace("温州葛瑞", "01部", inplace=True)

    df5["地区"].replace("台州唐惠", "台州1", inplace=True)
    df5["地区编码"].replace("台州唐惠", "0103", inplace=True)
    df5["负责人"].replace("台州唐惠", "唐惠", inplace=True)
    df5["部门"].replace("台州唐惠", "01部", inplace=True)

    df5["地区"].replace("温州潘磊", "温州2", inplace=True)
    df5["地区编码"].replace("温州潘磊", "0102", inplace=True)
    df5["负责人"].replace("温州潘磊", "潘磊", inplace=True)
    df5["部门"].replace("温州潘磊", "01部", inplace=True)

    df5["地区"].replace("台州胡文魁", "台州2", inplace=True)
    df5["地区编码"].replace("台州胡文魁", "0104", inplace=True)
    df5["负责人"].replace("台州胡文魁", "胡文魁", inplace=True)
    df5["部门"].replace("台州胡文魁", "01部", inplace=True)

    df5["地区"].replace("丽水", "丽水", inplace=True)
    df5["地区编码"].replace("丽水", "0105", inplace=True)
    df5["负责人"].replace("丽水", "方汝泼", inplace=True)
    df5["部门"].replace("丽水", "01部", inplace=True)


   #####一部完毕
    df5["地区"].replace("宁波市区", "宁波", inplace=True)
    df5["地区编码"].replace("宁波市区", "0201", inplace=True)
    df5["负责人"].replace("宁波市区", "丁玲", inplace=True)
    df5["部门"].replace("宁波市区", "02部", inplace=True)

    df5["地区"].replace("舟山北仑", "舟山北仑", inplace=True)
    df5["地区编码"].replace("舟山北仑", "0202", inplace=True)
    df5["负责人"].replace("舟山北仑", "高大勇", inplace=True)
    df5["部门"].replace("舟山北仑", "02部", inplace=True)

    df5["地区"].replace("慈溪余姚镇海", "北三县", inplace=True)
    df5["地区编码"].replace("慈溪余姚镇海", "0203", inplace=True)
    df5["负责人"].replace("慈溪余姚镇海", "陆金耀", inplace=True)
    df5["部门"].replace("慈溪余姚镇海", "02部", inplace=True)

    df5["地区"].replace("奉化宁海象山", "南三县", inplace=True)
    df5["地区编码"].replace("奉化宁海象山", "0204", inplace=True)
    df5["负责人"].replace("奉化宁海象山", "吴燕江", inplace=True)
    df5["部门"].replace("奉化宁海象山", "02部", inplace=True)
   ####三部####

    df5["地区"].replace("杭州姜立民", "省级", inplace=True)
    df5["地区编码"].replace("杭州姜立民", "0301", inplace=True)
    df5["负责人"].replace("杭州姜立民", "姜立民", inplace=True)
    df5["部门"].replace("杭州姜立民", "03部", inplace=True)

    df5["地区"].replace("杭州石亚国", "省级", inplace=True)
    df5["地区编码"].replace("杭州石亚国", "0301", inplace=True)
    df5["负责人"].replace("杭州石亚国", "姜立民", inplace=True)
    df5["部门"].replace("杭州石亚国", "03部", inplace=True)

    df5["地区"].replace("杭州陈靓", "省级", inplace=True)
    df5["地区编码"].replace("杭州陈靓", "0301", inplace=True)
    df5["负责人"].replace("杭州陈靓", "姜立民", inplace=True)
    df5["部门"].replace("杭州陈靓", "03部", inplace=True)

    df5["地区"].replace("杭州沈剑芳", "杭州", inplace=True)
    df5["地区编码"].replace("杭州沈剑芳", "0302", inplace=True)
    df5["负责人"].replace("杭州沈剑芳", "沈剑芳", inplace=True)
    df5["部门"].replace("杭州沈剑芳", "03部", inplace=True)

    df5["地区"].replace("杭州周海波", "杭州", inplace=True)
    df5["地区编码"].replace("杭州周海波", "0302", inplace=True)
    df5["负责人"].replace("杭州周海波", "沈剑芳", inplace=True)
    df5["部门"].replace("杭州周海波", "03部", inplace=True)

    df5["地区"].replace("嘉兴阮芳", "嘉湖", inplace=True)
    df5["地区编码"].replace("嘉兴阮芳", "0303", inplace=True)
    df5["负责人"].replace("嘉兴阮芳", "阮芳", inplace=True)
    df5["部门"].replace("嘉兴阮芳", "03部", inplace=True)

    df5["地区"].replace("湖州陈荣斌", "嘉湖", inplace=True)
    df5["地区编码"].replace("湖州陈荣斌", "0303", inplace=True)
    df5["负责人"].replace("湖州陈荣斌", "阮芳", inplace=True)
    df5["部门"].replace("湖州陈荣斌", "03部", inplace=True)

    df5["地区"].replace("晋江运城郑良", "晋江运城", inplace=True)
    df5["地区编码"].replace("晋江运城郑良", "0304", inplace=True)
    df5["负责人"].replace("晋江运城郑良", "郑良", inplace=True)
    df5["部门"].replace("晋江运城郑良", "03部", inplace=True)
  ####四部

    df5["地区"].replace("南京高跃", "南京1", inplace=True)
    df5["地区编码"].replace("南京高跃", "0401", inplace=True)
    df5["负责人"].replace("南京高跃", "高跃", inplace=True)
    df5["部门"].replace("南京高跃", "04部", inplace=True)

    df5["地区"].replace("南京阮建锋", "南京2", inplace=True)
    df5["地区编码"].replace("南京阮建锋", "0402", inplace=True)
    df5["负责人"].replace("南京阮建锋", "阮建锋", inplace=True)
    df5["部门"].replace("南京阮建锋", "04部", inplace=True)

    df5["地区"].replace("南京刘纪彬", "南京3", inplace=True)
    df5["地区编码"].replace("南京刘纪彬", "0403", inplace=True)
    df5["负责人"].replace("南京刘纪彬", "刘纪彬", inplace=True)
    df5["部门"].replace("南京刘纪彬", "04部", inplace=True)

    df5["地区"].replace("南京陈豪", "南京4", inplace=True)
    df5["地区编码"].replace("南京陈豪", "0404", inplace=True)
    df5["负责人"].replace("南京陈豪", "陈豪", inplace=True)
    df5["部门"].replace("南京陈豪", "04部", inplace=True)

  ###五部
    df5["地区"].replace("南通朱一亦", "南通1", inplace=True)
    df5["地区编码"].replace("南通朱一亦", "0501", inplace=True)
    df5["负责人"].replace("南通朱一亦", "李国旺 朱一亦", inplace=True)
    df5["部门"].replace("南通朱一亦", "05部", inplace=True)

    df5["地区"].replace("南通王峥骅", "南通2", inplace=True)
    df5["地区编码"].replace("南通王峥骅", "0502", inplace=True)
    df5["负责人"].replace("南通王峥骅", "李国旺 王铮骅", inplace=True)
    df5["部门"].replace("南通王峥骅", "05部", inplace=True)

    df5["地区"].replace("盐城", "盐城", inplace=True)
    df5["地区编码"].replace("盐城", "0503", inplace=True)
    df5["负责人"].replace("盐城", "岑潭泽 潘前进", inplace=True)
    df5["部门"].replace("盐城", "05部", inplace=True)

    df5["地区"].replace("连云港", "连云港", inplace=True)
    df5["地区编码"].replace("连云港", "0504", inplace=True)
    df5["负责人"].replace("连云港", "胡士艳 姜健", inplace=True)
    df5["部门"].replace("连云港", "05部", inplace=True)

  ###六部

    df5["地区"].replace("上海1", "上海1", inplace=True)
    df5["地区编码"].replace("上海1", "0601", inplace=True)
    df5["负责人"].replace("上海1", "汤俊", inplace=True)
    df5["部门"].replace("上海1", "06部", inplace=True)

    df5["地区"].replace("上海2", "上海2", inplace=True)
    df5["地区编码"].replace("上海2", "0602", inplace=True)
    df5["负责人"].replace("上海2", "邬幼波", inplace=True)
    df5["部门"].replace("上海2", "06部", inplace=True)

  ####七部

    df5["地区"].replace("苏州", "苏州", inplace=True)
    df5["地区编码"].replace("苏州", "0701", inplace=True)
    df5["负责人"].replace("苏州", "陈凯", inplace=True)
    df5["部门"].replace("苏州", "07部", inplace=True)

    df5["地区"].replace("苏州市郊", "苏郊", inplace=True)
    df5["地区编码"].replace("苏州市郊", "0702", inplace=True)
    df5["负责人"].replace("苏州市郊", "吕楠", inplace=True)
    df5["部门"].replace("苏州市郊", "07部", inplace=True)


   ####八部

    df5["地区"].replace("扬泰1", "扬泰1", inplace=True)
    df5["地区编码"].replace("扬泰1", "0801", inplace=True)
    df5["负责人"].replace("扬泰1", "姜海涛", inplace=True)
    df5["部门"].replace("扬泰1", "08部", inplace=True)

    df5["地区"].replace("扬泰2", "扬泰2", inplace=True)
    df5["地区编码"].replace("扬泰2", "0802", inplace=True)
    df5["负责人"].replace("扬泰2", "吕淳昱", inplace=True)
    df5["部门"].replace("扬泰2", "08部", inplace=True)

    df5["地区"].replace("扬泰3", "扬泰3", inplace=True)
    df5["地区编码"].replace("扬泰3", "0803", inplace=True)
    df5["负责人"].replace("扬泰3", "胡霞", inplace=True)
    df5["部门"].replace("扬泰3", "08部", inplace=True)


   ####九部

    df5["地区"].replace("徐州于博", "徐州", inplace=True)
    df5["地区编码"].replace("徐州于博", "0901", inplace=True)
    df5["负责人"].replace("徐州于博", "唐维洲 于博", inplace=True)
    df5["部门"].replace("徐州于博", "09部", inplace=True)

    df5["地区"].replace("徐州张浩", "徐州", inplace=True)
    df5["地区编码"].replace("徐州张浩", "0901", inplace=True)
    df5["负责人"].replace("徐州张浩", "唐维洲 于博", inplace=True)
    df5["部门"].replace("徐州张浩", "09部", inplace=True)

    df5["地区"].replace("宿迁", "宿迁", inplace=True)
    df5["地区编码"].replace("宿迁", "0902", inplace=True)
    df5["负责人"].replace("宿迁", "赵晨阳 王涛", inplace=True)
    df5["部门"].replace("宿迁", "09部", inplace=True)

    df5["地区"].replace("淮安", "淮安", inplace=True)
    df5["地区编码"].replace("淮安", "0903", inplace=True)
    df5["负责人"].replace("淮安", "赵晨阳 白虹", inplace=True)
    df5["部门"].replace("淮安", "09部", inplace=True)


    ####十部

    df5["地区"].replace("常州", "常州", inplace=True)
    df5["地区编码"].replace("常州", "1001", inplace=True)
    df5["负责人"].replace("常州", "吴羚", inplace=True)
    df5["部门"].replace("常州", "10部", inplace=True)

    df5["地区"].replace("镇江", "镇江", inplace=True)
    df5["地区编码"].replace("镇江", "1002", inplace=True)
    df5["负责人"].replace("镇江", "周丹", inplace=True)
    df5["部门"].replace("镇江", "10部", inplace=True)

    df5["地区"].replace("常镇", "常镇", inplace=True)
    df5["地区编码"].replace("常镇", "1003", inplace=True)
    df5["负责人"].replace("常镇", "于亚惠", inplace=True)
    df5["部门"].replace("常镇", "10部", inplace=True)

   ####十一部

    df5["地区"].replace("绍兴龚群波", "绍兴", inplace=True)
    df5["地区编码"].replace("绍兴龚群波", "1101", inplace=True)
    df5["负责人"].replace("绍兴龚群波", "龚群波", inplace=True)
    df5["部门"].replace("绍兴龚群波", "11部", inplace=True)

    df5["地区"].replace("衢州", "金衢", inplace=True)
    df5["地区编码"].replace("衢州", "1102", inplace=True)
    df5["负责人"].replace("衢州", "胡迪锋", inplace=True)
    df5["部门"].replace("衢州", "11部", inplace=True)

    df5["地区"].replace("金华", "金衢", inplace=True)
    df5["地区编码"].replace("金华", "1102", inplace=True)
    df5["负责人"].replace("金华", "胡迪锋", inplace=True)
    df5["部门"].replace("金华", "11部", inplace=True)

   ####十二部
    df5["地区"].replace("无锡裘涌", "无锡1", inplace=True)
    df5["地区编码"].replace("无锡裘涌", "1201", inplace=True)
    df5["负责人"].replace("无锡裘涌", "裘涌", inplace=True)
    df5["部门"].replace("无锡裘涌", "12部", inplace=True)

    df5["地区"].replace("无锡张立伟、赵飞", "无锡2", inplace=True)
    df5["地区编码"].replace("无锡张立伟、赵飞", "1202", inplace=True)
    df5["负责人"].replace("无锡张立伟、赵飞", "张立伟 赵飞", inplace=True)
    df5["部门"].replace("无锡张立伟、赵飞", "12部", inplace=True)

    ####调拨

    df5["地区"].replace("北京", "调拨", inplace=True)
    df5["地区编码"].replace("北京", "1301", inplace=True)
    df5["负责人"].replace("北京", "孙婷婷", inplace=True)
    df5["部门"].replace("北京", "调拨", inplace=True)

    df5["地区"].replace("生化分销", "调拨", inplace=True)
    df5["地区编码"].replace("生化分销", "1301", inplace=True)
    df5["负责人"].replace("生化分销", "孙婷婷", inplace=True)
    df5["部门"].replace("生化分销", "调拨", inplace=True)

    df5["地区"].replace("调拨", "调拨", inplace=True)
    df5["地区编码"].replace("调拨", "1301", inplace=True)
    df5["负责人"].replace("调拨", "孙婷婷", inplace=True)
    df5["部门"].replace("调拨", "调拨", inplace=True)

    df5["地区"].replace("维修部", "调拨", inplace=True)
    df5["地区编码"].replace("维修部", "1301", inplace=True)
    df5["负责人"].replace("维修部", "孙婷婷", inplace=True)
    df5["部门"].replace("维修部", "调拨", inplace=True)

    df6 = df5[df5["部门"] != "关联企业"]



    tkinter.messagebox.showinfo("提醒", "请选择英克销售发票明细（新试剂）源文件");
    # 加入开票明细

    df51 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

    df52 = df51.fillna(0)
    df53 = df52.drop(df52.index[[0, 1]], axis=0);

    df53["仪器开票合计"] = df53["03原厂仪器"] + df53["04采购平台仪器"] + df53["0501国产辅助配置"] + df53["0502流水线辅助配置"]

    df53["非仪器开票合计"] = df53["010101免疫（代理）"] + df53["010102特定蛋白（代理）"] + df53["010103血球（代理）"] + df53["010104普通生化（代理）"] \
                     + df53["010105AU生化（代理）"] + df53["010106利德曼生化（代理）"] + df53["010107尿液（代理）"] + df53["010109微生物（代理）"] + \
                     df53["010110索灵（代理BC）"] \
                     + df53["010111免疫（AMH）"] + df53["010201血凝（代理）"] + df53["0103lmmucor"] + df53["0104索灵"] + df53[
                         "010501质控试剂（代理）"] + df53["010502伯乐其它试剂"] \
                     + df53["010601BNP试剂（代理）"] + df53["010701血气（代理）"] + df53["0108苏医（代理BC血球质控）"] + df53["020101干式生化"] + df3[
                         "020102普通生化"] \
                     + df53["020103血气"] + df53["020104特殊生化"] + df53["020201血球"] + df53["020202血凝"] + df53["020203尿液"] + df53[
                         "020204血库"] + df53["020206体液"] \
                     + df53["020301发光"] + df53["020302特定蛋白"] + df53["020303酶免类"] + df53["020304其它免疫"] + df53["020305厦门万泰"] + \
                     df53["0204微生物"] + df53["0205药字号"] \
                     + df53["0206分子诊断"] + df53["0207病理科"] + df53["0208采购平台其它"] + df53["0209质控"] + df53["06软件"] + df53["07配件"] \
                     + df53["08其它业务"] + df53["0901基因试剂（自产）"] + df53["0902基因试剂（其它厂家）"] + df53["1101强盛生化"] + df53[
                         "1201沃文特免疫"] + df53["1202沃文特其他"] \
                     + df53["99其它"]

    df54 = df53.drop(["Unnamed: 1", "Unnamed: 2", "Unnamed: 3"], axis=1)  # 删列

    df55 = df54.groupby(["Unnamed: 0"], as_index=False)["非仪器开票合计", "仪器开票合计"].sum();

    df55["地区"] = df55["Unnamed: 0"]
    df55["负责人"] = df55["Unnamed: 0"]
    df55["地区编码"] = df55["Unnamed: 0"]
    df55["部门"] = df55["Unnamed: 0"]

    df55["地区"].replace("温州葛瑞", "温州1", inplace=True)
    df55["地区编码"].replace("温州葛瑞", "0101", inplace=True)
    df55["负责人"].replace("温州葛瑞", "葛瑞", inplace=True)
    df55["部门"].replace("温州葛瑞", "01部", inplace=True)

    df55["地区"].replace("台州唐惠", "台州1", inplace=True)
    df55["地区编码"].replace("台州唐惠", "0103", inplace=True)
    df55["负责人"].replace("台州唐惠", "唐惠", inplace=True)
    df55["部门"].replace("台州唐惠", "01部", inplace=True)

    df55["地区"].replace("温州潘磊", "温州2", inplace=True)
    df55["地区编码"].replace("温州潘磊", "0102", inplace=True)
    df55["负责人"].replace("温州潘磊", "潘磊", inplace=True)
    df55["部门"].replace("温州潘磊", "01部", inplace=True)

    df55["地区"].replace("台州胡文魁", "台州2", inplace=True)
    df55["地区编码"].replace("台州胡文魁", "0104", inplace=True)
    df55["负责人"].replace("台州胡文魁", "胡文魁", inplace=True)
    df55["部门"].replace("台州胡文魁", "01部", inplace=True)

    df55["地区"].replace("丽水", "丽水", inplace=True)
    df55["地区编码"].replace("丽水", "0105", inplace=True)
    df55["负责人"].replace("丽水", "方汝泼", inplace=True)
    df55["部门"].replace("丽水", "01部", inplace=True)

    #####一部完毕
    df55["地区"].replace("宁波市区", "宁波", inplace=True)
    df55["地区编码"].replace("宁波市区", "0201", inplace=True)
    df55["负责人"].replace("宁波市区", "丁玲", inplace=True)
    df55["部门"].replace("宁波市区", "02部", inplace=True)

    df55["地区"].replace("舟山北仑", "舟山北仑", inplace=True)
    df55["地区编码"].replace("舟山北仑", "0202", inplace=True)
    df55["负责人"].replace("舟山北仑", "高大勇", inplace=True)
    df55["部门"].replace("舟山北仑", "02部", inplace=True)

    df55["地区"].replace("慈溪余姚镇海", "北三县", inplace=True)
    df55["地区编码"].replace("慈溪余姚镇海", "0203", inplace=True)
    df55["负责人"].replace("慈溪余姚镇海", "陆金耀", inplace=True)
    df55["部门"].replace("慈溪余姚镇海", "02部", inplace=True)

    df55["地区"].replace("奉化宁海象山", "南三县", inplace=True)
    df55["地区编码"].replace("奉化宁海象山", "0204", inplace=True)
    df55["负责人"].replace("奉化宁海象山", "吴燕江", inplace=True)
    df55["部门"].replace("奉化宁海象山", "02部", inplace=True)
    ####三部####

    df55["地区"].replace("杭州姜立民", "省级", inplace=True)
    df55["地区编码"].replace("杭州姜立民", "0301", inplace=True)
    df55["负责人"].replace("杭州姜立民", "姜立民", inplace=True)
    df55["部门"].replace("杭州姜立民", "03部", inplace=True)

    df55["地区"].replace("杭州石亚国", "省级", inplace=True)
    df55["地区编码"].replace("杭州石亚国", "0301", inplace=True)
    df55["负责人"].replace("杭州石亚国", "姜立民", inplace=True)
    df55["部门"].replace("杭州石亚国", "03部", inplace=True)

    df55["地区"].replace("杭州陈靓", "省级", inplace=True)
    df55["地区编码"].replace("杭州陈靓", "0301", inplace=True)
    df55["负责人"].replace("杭州陈靓", "姜立民", inplace=True)
    df55["部门"].replace("杭州陈靓", "03部", inplace=True)

    df55["地区"].replace("杭州沈剑芳", "杭州", inplace=True)
    df55["地区编码"].replace("杭州沈剑芳", "0302", inplace=True)
    df55["负责人"].replace("杭州沈剑芳", "沈剑芳", inplace=True)
    df55["部门"].replace("杭州沈剑芳", "03部", inplace=True)

    df55["地区"].replace("杭州周海波", "杭州", inplace=True)
    df55["地区编码"].replace("杭州周海波", "0302", inplace=True)
    df55["负责人"].replace("杭州周海波", "沈剑芳", inplace=True)
    df55["部门"].replace("杭州周海波", "03部", inplace=True)

    df55["地区"].replace("嘉兴阮芳", "嘉湖", inplace=True)
    df55["地区编码"].replace("嘉兴阮芳", "0303", inplace=True)
    df55["负责人"].replace("嘉兴阮芳", "阮芳", inplace=True)
    df55["部门"].replace("嘉兴阮芳", "03部", inplace=True)

    df55["地区"].replace("湖州陈荣斌", "嘉湖", inplace=True)
    df55["地区编码"].replace("湖州陈荣斌", "0303", inplace=True)
    df55["负责人"].replace("湖州陈荣斌", "阮芳", inplace=True)
    df55["部门"].replace("湖州陈荣斌", "03部", inplace=True)

    df55["地区"].replace("晋江运城郑良", "晋江运城", inplace=True)
    df55["地区编码"].replace("晋江运城郑良", "0304", inplace=True)
    df55["负责人"].replace("晋江运城郑良", "郑良", inplace=True)
    df55["部门"].replace("晋江运城郑良", "03部", inplace=True)
    ####四部

    df55["地区"].replace("南京高跃", "南京1", inplace=True)
    df55["地区编码"].replace("南京高跃", "0401", inplace=True)
    df55["负责人"].replace("南京高跃", "高跃", inplace=True)
    df55["部门"].replace("南京高跃", "04部", inplace=True)

    df55["地区"].replace("南京阮建锋", "南京2", inplace=True)
    df55["地区编码"].replace("南京阮建锋", "0402", inplace=True)
    df55["负责人"].replace("南京阮建锋", "阮建锋", inplace=True)
    df55["部门"].replace("南京阮建锋", "04部", inplace=True)

    df55["地区"].replace("南京刘纪彬", "南京3", inplace=True)
    df55["地区编码"].replace("南京刘纪彬", "0403", inplace=True)
    df55["负责人"].replace("南京刘纪彬", "刘纪彬", inplace=True)
    df55["部门"].replace("南京刘纪彬", "04部", inplace=True)

    df55["地区"].replace("南京陈豪", "南京4", inplace=True)
    df55["地区编码"].replace("南京陈豪", "0404", inplace=True)
    df55["负责人"].replace("南京陈豪", "陈豪", inplace=True)
    df55["部门"].replace("南京陈豪", "04部", inplace=True)

    ###五部
    df55["地区"].replace("南通朱一亦", "南通1", inplace=True)
    df55["地区编码"].replace("南通朱一亦", "0501", inplace=True)
    df55["负责人"].replace("南通朱一亦", "李国旺 朱一亦", inplace=True)
    df55["部门"].replace("南通朱一亦", "05部", inplace=True)

    df55["地区"].replace("南通王峥骅", "南通2", inplace=True)
    df55["地区编码"].replace("南通王峥骅", "0502", inplace=True)
    df55["负责人"].replace("南通王峥骅", "李国旺 王铮骅", inplace=True)
    df55["部门"].replace("南通王峥骅", "05部", inplace=True)

    df55["地区"].replace("盐城", "盐城", inplace=True)
    df55["地区编码"].replace("盐城", "0503", inplace=True)
    df55["负责人"].replace("盐城", "岑潭泽 潘前进", inplace=True)
    df55["部门"].replace("盐城", "05部", inplace=True)

    df55["地区"].replace("连云港", "连云港", inplace=True)
    df55["地区编码"].replace("连云港", "0504", inplace=True)
    df55["负责人"].replace("连云港", "胡士艳 姜健", inplace=True)
    df55["部门"].replace("连云港", "05部", inplace=True)

    ###六部

    df55["地区"].replace("上海1", "上海1", inplace=True)
    df55["地区编码"].replace("上海1", "0601", inplace=True)
    df55["负责人"].replace("上海1", "汤俊", inplace=True)
    df55["部门"].replace("上海1", "06部", inplace=True)

    df55["地区"].replace("上海2", "上海2", inplace=True)
    df55["地区编码"].replace("上海2", "0602", inplace=True)
    df55["负责人"].replace("上海2", "邬幼波", inplace=True)
    df55["部门"].replace("上海2", "06部", inplace=True)

    ####七部

    df55["地区"].replace("苏州", "苏州", inplace=True)
    df55["地区编码"].replace("苏州", "0701", inplace=True)
    df55["负责人"].replace("苏州", "陈凯", inplace=True)
    df55["部门"].replace("苏州", "07部", inplace=True)

    df55["地区"].replace("苏州市郊", "苏郊", inplace=True)
    df55["地区编码"].replace("苏州市郊", "0702", inplace=True)
    df55["负责人"].replace("苏州市郊", "吕楠", inplace=True)
    df55["部门"].replace("苏州市郊", "07部", inplace=True)

    ####八部

    df55["地区"].replace("扬泰1", "扬泰1", inplace=True)
    df55["地区编码"].replace("扬泰1", "0801", inplace=True)
    df55["负责人"].replace("扬泰1", "姜海涛", inplace=True)
    df55["部门"].replace("扬泰1", "08部", inplace=True)

    df55["地区"].replace("扬泰2", "扬泰2", inplace=True)
    df55["地区编码"].replace("扬泰2", "0802", inplace=True)
    df55["负责人"].replace("扬泰2", "吕淳昱", inplace=True)
    df55["部门"].replace("扬泰2", "08部", inplace=True)

    df55["地区"].replace("扬泰3", "扬泰3", inplace=True)
    df55["地区编码"].replace("扬泰3", "0803", inplace=True)
    df55["负责人"].replace("扬泰3", "胡霞", inplace=True)
    df55["部门"].replace("扬泰3", "08部", inplace=True)

    ####九部

    df55["地区"].replace("徐州于博", "徐州", inplace=True)
    df55["地区编码"].replace("徐州于博", "0901", inplace=True)
    df55["负责人"].replace("徐州于博", "唐维洲 于博", inplace=True)
    df55["部门"].replace("徐州于博", "09部", inplace=True)

    df55["地区"].replace("徐州张浩", "徐州", inplace=True)
    df55["地区编码"].replace("徐州张浩", "0901", inplace=True)
    df55["负责人"].replace("徐州张浩", "唐维洲 于博", inplace=True)
    df55["部门"].replace("徐州张浩", "09部", inplace=True)

    df55["地区"].replace("宿迁", "宿迁", inplace=True)
    df55["地区编码"].replace("宿迁", "0902", inplace=True)
    df55["负责人"].replace("宿迁", "赵晨阳 王涛", inplace=True)
    df55["部门"].replace("宿迁", "09部", inplace=True)

    df55["地区"].replace("淮安", "淮安", inplace=True)
    df55["地区编码"].replace("淮安", "0903", inplace=True)
    df55["负责人"].replace("淮安", "赵晨阳 白虹", inplace=True)
    df55["部门"].replace("淮安", "09部", inplace=True)

    ####十部

    df55["地区"].replace("常州", "常州", inplace=True)
    df55["地区编码"].replace("常州", "1001", inplace=True)
    df55["负责人"].replace("常州", "吴羚", inplace=True)
    df55["部门"].replace("常州", "10部", inplace=True)

    df55["地区"].replace("镇江", "镇江", inplace=True)
    df55["地区编码"].replace("镇江", "1002", inplace=True)
    df55["负责人"].replace("镇江", "周丹", inplace=True)
    df55["部门"].replace("镇江", "10部", inplace=True)

    df55["地区"].replace("常镇", "常镇", inplace=True)
    df55["地区编码"].replace("常镇", "1003", inplace=True)
    df55["负责人"].replace("常镇", "于亚惠", inplace=True)
    df55["部门"].replace("常镇", "10部", inplace=True)

    ####十一部

    df55["地区"].replace("绍兴龚群波", "绍兴", inplace=True)
    df55["地区编码"].replace("绍兴龚群波", "1101", inplace=True)
    df55["负责人"].replace("绍兴龚群波", "龚群波", inplace=True)
    df55["部门"].replace("绍兴龚群波", "11部", inplace=True)

    df55["地区"].replace("衢州", "金衢", inplace=True)
    df55["地区编码"].replace("衢州", "1102", inplace=True)
    df55["负责人"].replace("衢州", "胡迪锋", inplace=True)
    df55["部门"].replace("衢州", "11部", inplace=True)

    df55["地区"].replace("金华", "金衢", inplace=True)
    df55["地区编码"].replace("金华", "1102", inplace=True)
    df55["负责人"].replace("金华", "胡迪锋", inplace=True)
    df55["部门"].replace("金华", "11部", inplace=True)

    ####十二部
    df55["地区"].replace("无锡裘涌", "无锡1", inplace=True)
    df55["地区编码"].replace("无锡裘涌", "1201", inplace=True)
    df55["负责人"].replace("无锡裘涌", "裘涌", inplace=True)
    df55["部门"].replace("无锡裘涌", "12部", inplace=True)

    df55["地区"].replace("无锡张立伟、赵飞", "无锡2", inplace=True)
    df55["地区编码"].replace("无锡张立伟、赵飞", "1202", inplace=True)
    df55["负责人"].replace("无锡张立伟、赵飞", "张立伟 赵飞", inplace=True)
    df55["部门"].replace("无锡张立伟、赵飞", "12部", inplace=True)

    ####调拨

    df55["地区"].replace("北京", "调拨", inplace=True)
    df55["地区编码"].replace("北京", "1301", inplace=True)
    df55["负责人"].replace("北京", "孙婷婷", inplace=True)
    df55["部门"].replace("北京", "调拨", inplace=True)

    df55["地区"].replace("生化分销", "调拨", inplace=True)
    df55["地区编码"].replace("生化分销", "1301", inplace=True)
    df55["负责人"].replace("生化分销", "孙婷婷", inplace=True)
    df55["部门"].replace("生化分销", "调拨", inplace=True)

    df55["地区"].replace("调拨", "调拨", inplace=True)
    df55["地区编码"].replace("调拨", "1301", inplace=True)
    df55["负责人"].replace("调拨", "孙婷婷", inplace=True)
    df55["部门"].replace("调拨", "调拨", inplace=True)

    df55["地区"].replace("维修部", "调拨", inplace=True)
    df55["地区编码"].replace("维修部", "1301", inplace=True)
    df55["负责人"].replace("维修部", "孙婷婷", inplace=True)
    df55["部门"].replace("维修部", "调拨", inplace=True)

    df56 = df55[df55["部门"] != "关联企业"]
    df57 = df56.drop(["部门","地区","负责人","Unnamed: 0"], axis=1)  # 删列



    df60 = pd.merge(df6, df57, how='left', on=['地区编码']);  # 完全相同合并

    df61 = df60.drop(["Unnamed: 0"], axis=1)  # 删列

    df62 = df61.sort_values(by=['地区编码'], axis=0, ascending=True)  # 行排序

    df63 = df62.groupby(["部门","地区","负责人","地区编码"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计",  "仪器开票合计"].sum();
   ########下面小计不参与，直接用空表插入
    #####一部行小计开始df65
    df64 = df63[df63["部门"] == "01部"]
    df65 = df64.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();

    df65["地区"] = df65["部门"]
    df65["负责人"] = df65["部门"]
    df65["地区编码"] = df65["部门"]

    df65["地区编码"].replace("01部", "0199", inplace=True)
    df65["地区"].replace("01部", "诊断一部", inplace=True)
    df65["负责人"].replace("01部", "郭德春", inplace=True)
    df65["部门"].replace("01部", "浙南", inplace=True)

    print(df65)
    #####二部小计df67

    df66 = df63[df63["部门"] == "02部"]
    df67 = df66.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();

    df67["地区"] = df67["部门"]
    df67["负责人"] = df67["部门"]
    df67["地区编码"] = df67["部门"]

    df67["地区编码"].replace("02部", "0299", inplace=True)
    df67["地区"].replace("02部", "诊断二部", inplace=True)
    df67["负责人"].replace("02部", "余顶峰", inplace=True)
    df67["部门"].replace("02部", "浙东", inplace=True)
    print(df67)
    #####三部df69
    df68 = df63[df63["部门"] == "03部"]
    df69 = df68.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();
    df69["地区"] = df69["部门"]
    df69["负责人"] = df69["部门"]
    df69["地区编码"] = df69["部门"]
    df69["地区编码"].replace("03部", "0399", inplace=True)
    df69["地区"].replace("03部", "诊断三部", inplace=True)
    df69["负责人"].replace("03部", "叶仲华", inplace=True)
    df69["部门"].replace("03部", "杭嘉湖", inplace=True)
    print(df69)
    ####四部df71
    df70 = df63[df63["部门"] == "04部"]
    df71 = df70.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();
    df71["地区"] = df71["部门"]
    df71["负责人"] = df71["部门"]
    df71["地区编码"] = df71["部门"]
    df71["地区编码"].replace("04部", "0499", inplace=True)
    df71["地区"].replace("04部", "诊断四部", inplace=True)
    df71["负责人"].replace("04部", "吴珏", inplace=True)
    df71["部门"].replace("04部", "南京", inplace=True)
    print(df71)
    ####五部df73
    df72 = df63[df63["部门"] == "05部"]
    df73 = df72.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();
    df73["地区"] = df73["部门"]
    df73["负责人"] = df73["部门"]
    df73["地区编码"] = df73["部门"]
    df73["地区编码"].replace("05部", "0599", inplace=True)
    df73["地区"].replace("05部", "诊断五部", inplace=True)
    df73["负责人"].replace("05部", "屈建", inplace=True)
    df73["部门"].replace("05部", "苏中", inplace=True)
    print(df73)
    ####六部df75
    df74 = df63[df63["部门"] == "06部"]
    df75 = df74.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();
    df75["地区"] = df75["部门"]
    df75["负责人"] = df75["部门"]
    df75["地区编码"] = df75["部门"]
    df75["地区编码"].replace("06部", "0699", inplace=True)
    df75["地区"].replace("06部", "诊断六部", inplace=True)
    df75["负责人"].replace("06部", "邬幼波", inplace=True)
    df75["部门"].replace("06部", "上海", inplace=True)
    print(df75)
    ####七部部df77
    df76 = df63[df63["部门"] == "07部"]
    df77 = df76.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();
    df77["地区"] = df77["部门"]
    df77["负责人"] = df77["部门"]
    df77["地区编码"] = df77["部门"]
    df77["地区编码"].replace("07部", "0799", inplace=True)
    df77["地区"].replace("07部", "诊断七部", inplace=True)
    df77["负责人"].replace("07部", "全英娜", inplace=True)
    df77["部门"].replace("07部", "苏州", inplace=True)
    print(df77)
    ####八部df79
    df78 = df63[df63["部门"] == "08部"]
    df79 = df78.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();
    df79["地区"] = df79["部门"]
    df79["负责人"] = df79["部门"]
    df79["地区编码"] = df79["部门"]
    df79["地区编码"].replace("08部", "0899", inplace=True)
    df79["地区"].replace("08部", "诊断八部", inplace=True)
    df79["负责人"].replace("08部", "金英明 邹海洵", inplace=True)
    df79["部门"].replace("08部", "扬泰", inplace=True)
    print(df79)
    ####九部df81
    df80 = df63[df63["部门"] == "09部"]
    df81 = df80.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();
    df81["地区"] = df81["部门"]
    df81["负责人"] = df81["部门"]
    df81["地区编码"] = df81["部门"]
    df81["地区编码"].replace("09部", "0999", inplace=True)
    df81["地区"].replace("09部", "诊断九部", inplace=True)
    df81["负责人"].replace("09部", "吴蓓", inplace=True)
    df81["部门"].replace("09部", "楚宿徐", inplace=True)
    print(df81)
    ####十部df83
    df82 = df63[df63["部门"] == "10部"]
    df83 = df82.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();
    df83["地区"] = df83["部门"]
    df83["负责人"] = df83["部门"]
    df83["地区编码"] = df83["部门"]
    df83["地区编码"].replace("10部", "1099", inplace=True)
    df83["地区"].replace("10部", "诊断十部", inplace=True)
    df83["负责人"].replace("10部", "梅晓虹", inplace=True)
    df83["部门"].replace("10部", "镇常", inplace=True)
    print(df83)
    ####十一部df85
    df84 = df63[df63["部门"] == "11部"]
    df85 = df84.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();
    df85["地区"] = df85["部门"]
    df85["负责人"] = df85["部门"]
    df85["地区编码"] = df85["部门"]
    df85["地区编码"].replace("11部", "1199", inplace=True)
    df85["地区"].replace("11部", "诊断十一部", inplace=True)
    df85["负责人"].replace("11部", "李征", inplace=True)
    df85["部门"].replace("11部", "金衢绍", inplace=True)
    print(df85)
    ####十二部df87
    df86 = df63[df63["部门"] == "12部"]
    df87 = df86.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();
    df87["地区"] = df87["部门"]
    df87["负责人"] = df87["部门"]
    df87["地区编码"] = df87["部门"]
    df87["地区编码"].replace("12部", "1299", inplace=True)
    df87["地区"].replace("12部", "诊断十二部", inplace=True)
    df87["负责人"].replace("12部", "胡瑜 郭兵", inplace=True)
    df87["部门"].replace("12部", "无锡", inplace=True)
    print(df87)

    ####调拨df89
    df88 = df63[df63["部门"] == "调拨"]
    df89 = df88.groupby(["部门"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();
    df89["地区"] = df89["部门"]
    df89["负责人"] = df89["部门"]
    df89["地区编码"] = df89["部门"]
    df89["地区编码"].replace("调拨", "1399", inplace=True)
    df89["地区"].replace("调拨", "调拨", inplace=True)
    df89["负责人"].replace("调拨", "孙婷婷", inplace=True)
    df89["部门"].replace("调拨", "调拨", inplace=True)
    print(df89)



    ########总计
    df100 = pd.concat([df89,df87, df85, df83, df81, df79, df77, df75, df73, df71, df69,
                       df67, df65],
                      ignore_index=True)  # 组合
    df100["地区编码"].replace("1399", "9999", inplace=True)
    df100["地区编码"].replace("1299", "9999", inplace=True)
    df100["地区编码"].replace("1199", "9999", inplace=True)
    df100["地区编码"].replace("1099", "9999", inplace=True)
    df100["地区编码"].replace("0999", "9999", inplace=True)
    df100["地区编码"].replace("0899", "9999", inplace=True)
    df100["地区编码"].replace("0799", "9999", inplace=True)
    df100["地区编码"].replace("0699", "9999", inplace=True)
    df100["地区编码"].replace("0599", "9999", inplace=True)
    df100["地区编码"].replace("0499", "9999", inplace=True)
    df100["地区编码"].replace("0399", "9999", inplace=True)
    df100["地区编码"].replace("0299", "9999", inplace=True)
    df100["地区编码"].replace("0199", "9999", inplace=True)




    df101 = df100.groupby(["地区编码" ], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();

    df101["地区"] = df101["地区编码"]
    df101["负责人"] = df101["地区编码"]
    df101["部门"] = df101["地区编码"]
    df101["地区编码"].replace("9999", "9999", inplace=True)
    df101["地区"].replace("9999", "合计", inplace=True)
    df101["负责人"].replace("9999", "合计", inplace=True)
    df101["部门"].replace("9999", "合计", inplace=True)
    #####df101为大合计

    ###########大小合计组合

    df102 = pd.concat([df63,df101,df87, df85, df83, df81, df79, df77, df75, df73, df71, df69,
                       df67, df65],
                     ignore_index=True)  # 组合

    df103 = df102.sort_values(by=['地区编码'], axis=0, ascending=True)  # 行排序

    df104=df103.groupby(["部门","地区","负责人","地区编码"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();

    df105 = df104.sort_values(by=['地区编码'], axis=0, ascending=True)  # 行排序

   ########################上面小计暂时不参与

    ###插入虚拟表合计空白用于链接
   # df200=pd.dataframe([["浙南","诊断一部","郭德春","0199"],

    #                   columns=["部门","地区","负责人","地区编码"])
    data ={"地区编码":["0199","0299","0399","0499","0599","0699","0799","0899","0999","1099","1199","1299","9999"]}
    data1 ={"部门":["01部","02部","03部","04部","05部","06部","07部","08部","09部","10部","11部","12部","调拨"]}
    data2 = {"地区":["诊断一部","诊断二部","诊断三部","诊断四部","诊断五部","诊断六部","诊断七部","诊断八部","诊断九部","诊断十一部","诊断十二","调拨"]}
    data3 = {"负责人":["郭德春","余顶峰","叶仲华","吴珏","屈建","邬幼波","全英娜","金英明 邹海洵","吴蓓","梅晓虹","李征","胡瑜 郭兵","孙婷婷"]}

    df200 = pd.DataFrame(data)
    df201 = pd.DataFrame(data1)
    df202 = pd.DataFrame(data2)
    df203 = pd.DataFrame(data3)
    #"部门":["01部","02部","03部","04部","05部","06部","07部","08部","09部","10部","11部","12部","调拨"],"地区":["诊断一部","诊断二部","诊断三部","诊断四部","诊断五部","诊断六部","诊断七部","诊断八部","诊断九部","诊断十一部","诊断十二","调拨"],"负责人":["郭德春","余顶峰","叶仲华","吴珏","屈建","邬幼波","全英娜","金英明 邹海洵","吴蓓","梅晓虹","李征","胡瑜 郭兵","孙婷婷"],

    df204=pd.concat([df63,df200], ignore_index=True)  # 组合
    print(df204)

    df205 = df204.sort_values(by=['地区编码'], axis=0, ascending=True)  # 行排序
   ####1部207
    df206 = df205[df205["地区编码"] == "0199"]
    df207 =df206.fillna(0)

    df207["地区"].replace(0, "诊断一部", inplace=True)
    df207["负责人"].replace(0, "郭德春", inplace=True)
    df207["部门"].replace(0, "浙南", inplace=True)

    df207["非仪器出库合计"].replace(0, " ", inplace=True)
    df207["仪器出库合计"].replace(0, " ", inplace=True)
    df207["非仪器开票合计"].replace(0, " ", inplace=True)
    df207["仪器开票合计"].replace(0, " ", inplace=True)
   ###2部209
    df208 = df205[df205["地区编码"] == "0299"]
    df209 = df208.fillna(0)

    df209["地区"].replace(0, "诊断二部", inplace=True)
    df209["负责人"].replace(0, "余顶峰", inplace=True)
    df209["部门"].replace(0, "浙东", inplace=True)

    df209["非仪器出库合计"].replace(0, " ", inplace=True)
    df209["仪器出库合计"].replace(0, " ", inplace=True)
    df209["非仪器开票合计"].replace(0, " ", inplace=True)
    df209["仪器开票合计"].replace(0, " ", inplace=True)
   ####3部211
    df210 = df205[df205["地区编码"] == "0399"]
    df211 = df210.fillna(0)

    df211["地区"].replace(0, "诊断三部", inplace=True)
    df211["负责人"].replace(0, "毛存亮", inplace=True)
    df211["部门"].replace(0, "杭嘉湖", inplace=True)

    df211["非仪器出库合计"].replace(0, " ", inplace=True)
    df211["仪器出库合计"].replace(0, " ", inplace=True)
    df211["非仪器开票合计"].replace(0, " ", inplace=True)
    df211["仪器开票合计"].replace(0, " ", inplace=True)
  #####4部213
    df212 = df205[df205["地区编码"] == "0499"]
    df213 = df212.fillna(0)

    df213["地区"].replace(0, "诊断四部", inplace=True)
    df213["负责人"].replace(0, "吴珏", inplace=True)
    df213["部门"].replace(0, "南京", inplace=True)

    df213["非仪器出库合计"].replace(0, " ", inplace=True)
    df213["仪器出库合计"].replace(0, " ", inplace=True)
    df213["非仪器开票合计"].replace(0, " ", inplace=True)
    df213["仪器开票合计"].replace(0, " ", inplace=True)
    #####5部215
    df214 = df205[df205["地区编码"] == "0599"]
    df215 = df214.fillna(0)

    df215["地区"].replace(0, "诊断五部", inplace=True)
    df215["负责人"].replace(0, "屈建", inplace=True)
    df215["部门"].replace(0, "苏中", inplace=True)

    df215["非仪器出库合计"].replace(0, " ", inplace=True)
    df215["仪器出库合计"].replace(0, " ", inplace=True)
    df215["非仪器开票合计"].replace(0, " ", inplace=True)
    df215["仪器开票合计"].replace(0, " ", inplace=True)
    #####6部217
    df216 = df205[df205["地区编码"] == "0699"]
    df217 = df216.fillna(0)

    df217["地区"].replace(0, "诊断六部", inplace=True)
    df217["负责人"].replace(0, "邬幼波", inplace=True)
    df217["部门"].replace(0, "上海", inplace=True)

    df217["非仪器出库合计"].replace(0, " ", inplace=True)
    df217["仪器出库合计"].replace(0, " ", inplace=True)
    df217["非仪器开票合计"].replace(0, " ", inplace=True)
    df217["仪器开票合计"].replace(0, " ", inplace=True)

    #####7部219
    df218 = df205[df205["地区编码"] == "0799"]
    df219 = df218.fillna(0)

    df219["地区"].replace(0, "诊断七部", inplace=True)
    df219["负责人"].replace(0, "全英娜", inplace=True)
    df219["部门"].replace(0, "苏州", inplace=True)

    df219["非仪器出库合计"].replace(0, " ", inplace=True)
    df219["仪器出库合计"].replace(0, " ", inplace=True)
    df219["非仪器开票合计"].replace(0, " ", inplace=True)
    df219["仪器开票合计"].replace(0, " ", inplace=True)

    #####8部221
    df220 = df205[df205["地区编码"] == "0899"]
    df221 = df220.fillna(0)

    df221["地区"].replace(0, "诊断八部", inplace=True)
    df221["负责人"].replace(0, "金英明 邹海洵", inplace=True)
    df221["部门"].replace(0, "扬泰", inplace=True)

    df221["非仪器出库合计"].replace(0, " ", inplace=True)
    df221["仪器出库合计"].replace(0, " ", inplace=True)
    df221["非仪器开票合计"].replace(0, " ", inplace=True)
    df221["仪器开票合计"].replace(0, " ", inplace=True)

    #####9部223
    df222 = df205[df205["地区编码"] == "0999"]
    df223 = df222.fillna(0)

    df223["地区"].replace(0, "诊断九部", inplace=True)
    df223["负责人"].replace(0, "吴蓓", inplace=True)
    df223["部门"].replace(0, "楚宿徐", inplace=True)

    df223["非仪器出库合计"].replace(0, " ", inplace=True)
    df223["仪器出库合计"].replace(0, " ", inplace=True)
    df223["非仪器开票合计"].replace(0, " ", inplace=True)
    df223["仪器开票合计"].replace(0, " ", inplace=True)

    #####10部225
    df224 = df205[df205["地区编码"] == "1099"]
    df225 = df224.fillna(0)

    df225["地区"].replace(0, "诊断十部", inplace=True)
    df225["负责人"].replace(0, "梅晓虹", inplace=True)
    df225["部门"].replace(0, "镇常", inplace=True)

    df225["非仪器出库合计"].replace(0, " ", inplace=True)
    df225["仪器出库合计"].replace(0, " ", inplace=True)
    df225["非仪器开票合计"].replace(0, " ", inplace=True)
    df225["仪器开票合计"].replace(0, " ", inplace=True)

    #####11部227
    df226 = df205[df205["地区编码"] == "1199"]
    df227 = df226.fillna(0)

    df227["地区"].replace(0, "诊断十一部", inplace=True)
    df227["负责人"].replace(0, "李征", inplace=True)
    df227["部门"].replace(0, "金衢绍", inplace=True)

    df227["非仪器出库合计"].replace(0, " ", inplace=True)
    df227["仪器出库合计"].replace(0, " ", inplace=True)
    df227["非仪器开票合计"].replace(0, " ", inplace=True)
    df227["仪器开票合计"].replace(0, " ", inplace=True)

    #####12部229
    df228 = df205[df205["地区编码"] == "1299"]
    df229 = df228.fillna(0)

    df229["地区"].replace(0, "诊断十二部", inplace=True)
    df229["负责人"].replace(0, "胡瑜 郭兵", inplace=True)
    df229["部门"].replace(0, "无锡", inplace=True)

    df229["非仪器出库合计"].replace(0, " ", inplace=True)
    df229["仪器出库合计"].replace(0, " ", inplace=True)
    df229["非仪器开票合计"].replace(0, " ", inplace=True)
    df229["仪器开票合计"].replace(0, " ", inplace=True)



    #####总计233
    df232 = df205[df205["地区编码"] == "9999"]
    df233 = df232.fillna(0)

    df233["地区"].replace(0, "合计", inplace=True)
    df233["负责人"].replace(0, "合计", inplace=True)
    df233["部门"].replace(0, "合计", inplace=True)

    df233["非仪器出库合计"].replace(0, " ", inplace=True)
    df233["仪器出库合计"].replace(0, " ", inplace=True)
    df233["非仪器开票合计"].replace(0, " ", inplace=True)
    df233["仪器开票合计"].replace(0, " ", inplace=True)





    df235 = pd.concat([df63, df207,df209,df211,df213,df215,df217,df219,df221,df223,df225,df227,df229,df233], ignore_index=True)  # 组合

    df236=df235.groupby(["部门", "地区", "负责人", "地区编码"], as_index=False)[
        "非仪器出库合计", "仪器出库合计", "非仪器开票合计", "仪器开票合计"].sum();

    df237 = df236.sort_values(by=['地区编码'], axis=0, ascending=True)  # 行排序



    #df3 = df2.drop(df2.index[[0, 1]], axis=0);



    #df2['开票合计'] = df2.apply(lambda x: x.sum(), axis=1)








    df237.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                 filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                            ("Microsoft Excel 97-20003 文件", "*.xls")],
                                                                 defaultextension=".xlsx"));
    #df100.to_excel("客户等级测试" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xls", sheet_name="sheet1",
     #          index=False)  # 自动输出
    tkinter.messagebox.showinfo("运行结果", "销售开票对比导出成功！");

 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
             message="请检查提交源文件是否正确 '" + str(error) + "'.",
             detail=traceback.format_exc())









def appendStr28():  #######宁波医药回款天数计算

 try:
   tkinter.messagebox.showinfo("提醒", "请先选择金蝶辅助项目余额表源文件");
   df1 = pd.read_excel(tkinter.filedialog.askopenfilename());
   df2 = df1.fillna(0)
   df3 = df2.drop(df2.columns[[[ 7,9,15]]], axis=1)

   df4 = df3['核算项目编码'].str.split('-| ', expand=True);


   print(df4)
   df5 = pd.merge(df4, df3, right_index=True, left_index=True);#横向拼接
   df6 = df5.drop(df5.columns[[[[[[[[[[0,1,2,3,5,6,7,9,12,13]]]]]]]]]], axis=1)
   df7 = df6.rename(columns={'4': '业务员', 'Unnamed: 8': '年初余额', 'Unnamed: 10': '期初余额', 'Unnamed: 16': '期末余额'});
   df7["业务员"] = df7[4]
   #df8=df7.drop(df7.columns[0], axis=1)
   #df8 = df7.drop([4], axis=1)  # 删列
   df8=df7[df7["业务员"] != ""]
   df9 = df8.groupby(["业务员", "期间"], as_index=False)["年初余额","期初余额","借方","贷方","本年累计借方","本年累计贷方","期末余额"].sum();
   ####df9基准表
   ####下列表依据df9表进行组合
   ####第一部分 应收账款余额




   df10 = df9[df9["期间"] == 1]    #df63[df63["部门"] == "11部"]
   df11 = df10[df10["业务员"] == "傅诗云"]
   df12 = df11.drop(["期间","期初余额","借方","贷方","本年累计借方","本年累计贷方"], axis=1)  # 删列
   df13 = df12.rename(columns={'期末余额': '1月应收余额'});
   print(df13)
   df14 = df9[df9["期间"] == 2]
   df15 = df14[df14["业务员"] == "傅诗云"]
   df16 = df15.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方","年初余额"], axis=1)  # 删列
   df17 = df16.rename(columns={'期末余额': '2月应收余额'});
   print(df17)

   df18 = df9[df9["期间"] == 3]
   df19 = df18[df18["业务员"] == "傅诗云"]
   df20 = df19.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   df21 = df20.rename(columns={'期末余额': '3月应收余额'});
   print(df21)

   df22 = df9[df9["期间"] == 4]
   df23 = df22[df22["业务员"] == "傅诗云"]
   df24 = df23.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   df25 = df24.rename(columns={'期末余额': '4月应收余额'});
   print(df25)

   df26 = df9[df9["期间"] == 5]
   df27 = df26[df26["业务员"] == "傅诗云"]
   df28 = df27.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   df29 = df28.rename(columns={'期末余额': '5月应收余额'});
   print(df29)

   df30 = df9[df9["期间"] == 6]
   df31 = df30[df30["业务员"] == "傅诗云"]
   df32 = df31.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   df33 = df32.rename(columns={'期末余额': '6月应收余额'});
   print(df33)

   df34 = df9[df9["期间"] == 7]
   df35 = df34[df34["业务员"] == "傅诗云"]
   df36 = df35.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   df37 = df36.rename(columns={'期末余额': '7月应收余额'});
   print(df37)

   df38 = df9[df9["期间"] == 8]
   df39 = df38[df38["业务员"] == "傅诗云"]
   df40 = df39.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   df41 = df40.rename(columns={'期末余额': '8月应收余额'});
   print(df41)

   df42 = df9[df9["期间"] == 9]
   df43 = df42[df42["业务员"] == "傅诗云"]
   df44 = df43.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   df45 = df44.rename(columns={'期末余额': '9月应收余额'});
   print(df45)

   df46 = df9[df9["期间"] == 10]
   df47 = df46[df46["业务员"] == "傅诗云"]
   df48 = df47.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   df49 = df48.rename(columns={'期末余额': '10月应收余额'});
   print(df49)

   df50 = df9[df9["期间"] == 11]
   df51 = df50[df50["业务员"] == "傅诗云"]
   df52 = df51.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   df53 = df52.rename(columns={'期末余额': '11月应收余额'});
   print(df53)

   df54 = df9[df9["期间"] == 12]
   df55 = df54[df54["业务员"] == "傅诗云"]
   df56 = df55.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   df57 = df56.rename(columns={'期末余额': '12月应收余额'});
   print(df57)





   #df100 = pd.merge(df12, df15, right_index=True, left_index=True);
   df99=pd.merge(df13, df17,how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   df100 = pd.merge(df99, df21,how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   df101 = pd.merge(df100, df25, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   df102 = pd.merge(df101, df29, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   df103 = pd.merge(df102, df33, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   df104 = pd.merge(df103, df37, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   df105 = pd.merge(df104, df41, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   df106 = pd.merge(df105, df45, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   df107 = pd.merge(df106, df49, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   df108 = pd.merge(df107, df53, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   df109 = pd.merge(df108, df57, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)

   print(df109)
#######傅诗云完毕,戚肆朝开始a
   dfa10 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   dfa11 = dfa10[dfa10["业务员"] == "戚肆朝"]
   dfa12 = dfa11.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方"], axis=1)  # 删列
   dfa13 = dfa12.rename(columns={'期末余额': '1月应收余额'});
   print(dfa13)
   dfa14 = df9[df9["期间"] == 2]
   dfa15 = dfa14[dfa14["业务员"] == "戚肆朝"]
   dfa16 = dfa15.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfa17 = dfa16.rename(columns={'期末余额': '2月应收余额'});
   print(dfa17)

   dfa18 = df9[df9["期间"] == 3]
   dfa19 = dfa18[dfa18["业务员"] == "戚肆朝"]
   dfa20 = dfa19.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfa21 = dfa20.rename(columns={'期末余额': '3月应收余额'});
   print(dfa21)

   dfa22 = df9[df9["期间"] == 4]
   dfa23 = dfa22[dfa22["业务员"] == "戚肆朝"]
   dfa24 = dfa23.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfa25 = dfa24.rename(columns={'期末余额': '4月应收余额'});
   print(dfa25)

   dfa26 = df9[df9["期间"] == 5]
   dfa27 = dfa26[dfa26["业务员"] == "戚肆朝"]
   dfa28 = dfa27.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfa29 = dfa28.rename(columns={'期末余额': '5月应收余额'});
   print(dfa29)

   dfa30 = df9[df9["期间"] == 6]
   dfa31 = dfa30[dfa30["业务员"] == "戚肆朝"]
   dfa32 = dfa31.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfa33 = dfa32.rename(columns={'期末余额': '6月应收余额'});
   print(dfa33)

   dfa34 = df9[df9["期间"] == 7]
   dfa35 = dfa34[dfa34["业务员"] == "戚肆朝"]
   dfa36 = dfa35.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfa37 = dfa36.rename(columns={'期末余额': '7月应收余额'});
   print(dfa37)

   dfa38 = df9[df9["期间"] == 8]
   dfa39 = dfa38[dfa38["业务员"] == "戚肆朝"]
   dfa40 = dfa39.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfa41 = dfa40.rename(columns={'期末余额': '8月应收余额'});
   print(dfa41)

   dfa42 = df9[df9["期间"] == 9]
   dfa43 = dfa42[dfa42["业务员"] == "戚肆朝"]
   dfa44 = dfa43.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfa45 = dfa44.rename(columns={'期末余额': '9月应收余额'});
   print(dfa45)

   dfa46 = df9[df9["期间"] == 10]
   dfa47 = dfa46[dfa46["业务员"] == "戚肆朝"]
   dfa48 = dfa47.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfa49 = dfa48.rename(columns={'期末余额': '10月应收余额'});
   print(df49)

   dfa50 = df9[df9["期间"] == 11]
   dfa51 = dfa50[dfa50["业务员"] == "戚肆朝"]
   dfa52 = dfa51.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfa53 = dfa52.rename(columns={'期末余额': '11月应收余额'});
   print(dfa53)

   dfa54 = df9[df9["期间"] == 12]
   dfa55 = dfa54[dfa54["业务员"] == "戚肆朝"]
   dfa56 = dfa55.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfa57 = dfa56.rename(columns={'期末余额': '12月应收余额'});
   print(dfa57)

   # df100 = pd.merge(df12, df15, right_index=True, left_index=True);
   dfa99 = pd.merge(dfa13, dfa17, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfa100 = pd.merge(dfa99, dfa21, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfa101 = pd.merge(dfa100, dfa25, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfa102 = pd.merge(dfa101, dfa29, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfa103 = pd.merge(dfa102, dfa33, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfa104 = pd.merge(dfa103, dfa37, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfa105 = pd.merge(dfa104, dfa41, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfa106 = pd.merge(dfa105, dfa45, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfa107 = pd.merge(dfa106, dfa49, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfa108 = pd.merge(dfa107, dfa53, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfa109 = pd.merge(dfa108, dfa57, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)

   print(dfa109)

####张丽君开始b
   dfb10 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   dfb11 = dfb10[dfb10["业务员"] == "张丽君"]
   dfb12 = dfb11.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方"], axis=1)  # 删列
   dfb13 = dfb12.rename(columns={'期末余额': '1月应收余额'});
   print(dfb13)
   dfb14 = df9[df9["期间"] == 2]
   dfb15 = dfb14[dfb14["业务员"] == "张丽君"]
   dfb16 = dfb15.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfb17 = dfb16.rename(columns={'期末余额': '2月应收余额'});
   print(dfb17)

   dfb18 = df9[df9["期间"] == 3]
   dfb19 = dfb18[dfb18["业务员"] == "张丽君"]
   dfb20 = dfb19.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfb21 = dfb20.rename(columns={'期末余额': '3月应收余额'});
   print(dfb21)

   dfb22 = df9[df9["期间"] == 4]
   dfb23 = dfb22[dfb22["业务员"] == "张丽君"]
   dfb24 = dfb23.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfb25 = dfb24.rename(columns={'期末余额': '4月应收余额'});
   print(dfb25)

   dfb26 = df9[df9["期间"] == 5]
   dfb27 = dfb26[dfa26["业务员"] == "张丽君"]
   dfb28 = dfb27.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfb29 = dfb28.rename(columns={'期末余额': '5月应收余额'});
   print(dfb29)

   dfb30 = df9[df9["期间"] == 6]
   dfb31 = dfb30[dfb30["业务员"] == "张丽君"]
   dfb32 = dfb31.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfb33 = dfb32.rename(columns={'期末余额': '6月应收余额'});
   print(dfb33)

   dfb34 = df9[df9["期间"] == 7]
   dfb35 = dfb34[dfb34["业务员"] == "张丽君"]
   dfb36 = dfb35.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfb37 = dfb36.rename(columns={'期末余额': '7月应收余额'});
   print(dfb37)

   dfb38 = df9[df9["期间"] == 8]
   dfb39 = dfb38[dfb38["业务员"] == "张丽君"]
   dfb40 = dfb39.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfb41 = dfb40.rename(columns={'期末余额': '8月应收余额'});
   print(dfb41)

   dfb42 = df9[df9["期间"] == 9]
   dfb43 = dfb42[dfb42["业务员"] == "张丽君"]
   dfb44 = dfb43.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfb45 = dfb44.rename(columns={'期末余额': '9月应收余额'});
   print(dfb45)

   dfb46 = df9[df9["期间"] == 10]
   dfb47 = dfb46[dfb46["业务员"] == "张丽君"]
   dfb48 = dfb47.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfb49 = dfb48.rename(columns={'期末余额': '10月应收余额'});
   print(dfb49)

   dfb50 = df9[df9["期间"] == 11]
   dfb51 = dfb50[dfb50["业务员"] == "张丽君"]
   dfb52 = dfb51.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfb53 = dfb52.rename(columns={'期末余额': '11月应收余额'});
   print(dfb53)

   dfb54 = df9[df9["期间"] == 12]
   dfb55 = dfb54[dfb54["业务员"] == "张丽君"]
   dfb56 = dfb55.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfb57 = dfb56.rename(columns={'期末余额': '12月应收余额'});
   print(dfb57)



   # df100 = pd.merge(df12, df15, right_index=True, left_index=True);
   dfb99 = pd.merge(dfb13, dfb17, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfb100 = pd.merge(dfb99, dfb21, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfb101 = pd.merge(dfb100, dfb25, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfb102 = pd.merge(dfb101, dfb29, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfb103 = pd.merge(dfb102, dfb33, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfb104 = pd.merge(dfb103, dfb37, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfb105 = pd.merge(dfb104, dfb41, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfb106 = pd.merge(dfb105, dfb45, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfb107 = pd.merge(dfb106, dfb49, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfb108 = pd.merge(dfb107, dfb53, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfb109 = pd.merge(dfb108, dfb57, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)

   print(dfb109)

####王振杰开始c
   dfc10 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   dfc11 = dfc10[dfc10["业务员"] == "王振杰"]
   dfc12 = dfc11.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方"], axis=1)  # 删列
   dfc13 = dfc12.rename(columns={'期末余额': '1月应收余额'});
   print(dfc13)
   dfc14 = df9[df9["期间"] == 2]
   dfc15 = dfc14[dfc14["业务员"] == "王振杰"]
   dfc16 = dfc15.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfc17 = dfc16.rename(columns={'期末余额': '2月应收余额'});
   print(dfc17)

   dfc18 = df9[df9["期间"] == 3]
   dfc19 = dfc18[dfc18["业务员"] == "王振杰"]
   dfc20 = dfc19.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfc21 = dfc20.rename(columns={'期末余额': '3月应收余额'});
   print(dfc21)

   dfc22 = df9[df9["期间"] == 4]
   dfc23 = dfc22[dfc22["业务员"] == "王振杰"]
   dfc24 = dfc23.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfc25 = dfc24.rename(columns={'期末余额': '4月应收余额'});
   print(dfc25)

   dfc26 = df9[df9["期间"] == 5]
   dfc27 = dfc26[dfa26["业务员"] == "王振杰"]
   dfc28 = dfc27.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfc29 = dfc28.rename(columns={'期末余额': '5月应收余额'});
   print(dfc29)

   dfc30 = df9[df9["期间"] == 6]
   dfc31 = dfc30[dfc30["业务员"] == "王振杰"]
   dfc32 = dfc31.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfc33 = dfc32.rename(columns={'期末余额': '6月应收余额'});
   print(dfc33)

   dfc34 = df9[df9["期间"] == 7]
   dfc35 = dfc34[dfc34["业务员"] == "王振杰"]
   dfc36 = dfc35.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfc37 = dfc36.rename(columns={'期末余额': '7月应收余额'});
   print(dfc37)

   dfc38 = df9[df9["期间"] == 8]
   dfc39 = dfc38[dfc38["业务员"] == "王振杰"]
   dfc40 = dfc39.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfc41 = dfc40.rename(columns={'期末余额': '8月应收余额'});
   print(dfc41)

   dfc42 = df9[df9["期间"] == 9]
   dfc43 = dfc42[dfc42["业务员"] == "王振杰"]
   dfc44 = dfc43.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfc45 = dfc44.rename(columns={'期末余额': '9月应收余额'});
   print(dfc45)

   dfc46 = df9[df9["期间"] == 10]
   dfc47 = dfc46[dfc46["业务员"] == "王振杰"]
   dfc48 = dfc47.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfc49 = dfc48.rename(columns={'期末余额': '10月应收余额'});
   print(dfc49)

   dfc50 = df9[df9["期间"] == 11]
   dfc51 = dfc50[dfc50["业务员"] == "王振杰"]
   dfc52 = dfc51.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfc53 = dfc52.rename(columns={'期末余额': '11月应收余额'});
   print(dfc53)

   dfc54 = df9[df9["期间"] == 12]
   dfc55 = dfc54[dfc54["业务员"] == "王振杰"]
   dfc56 = dfc55.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfc57 = dfc56.rename(columns={'期末余额': '12月应收余额'});
   print(dfc57)

   # df100 = pd.merge(df12, df15, right_index=True, left_index=True);
   dfc99 = pd.merge(dfc13, dfc17, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfc100 = pd.merge(dfc99, dfc21, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfc101 = pd.merge(dfc100, dfc25, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfc102 = pd.merge(dfc101, dfc29, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfc103 = pd.merge(dfc102, dfc33, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfc104 = pd.merge(dfc103, dfc37, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfc105 = pd.merge(dfc104, dfc41, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfc106 = pd.merge(dfc105, dfc45, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfc107 = pd.merge(dfc106, dfc49, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfc108 = pd.merge(dfc107, dfc53, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfc109 = pd.merge(dfc108, dfc57, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)

   print(dfc109)


   df1000=pd.concat([df109,dfa109,dfb109,dfc109],ignore_index=True)  # 组合
   ####药品一部合计
   df1000["部门"]=df1000["业务员"]
   #df1001=df1000.rename(columns={'傅诗云': '药品1部','戚肆朝': '药品1部','张丽君': '药品1部','王振杰': '药品1部'});

   df1000["部门"].replace("傅诗云", "药品1部", inplace=True)
   df1000["部门"].replace("戚肆朝", "药品1部", inplace=True)
   df1000["部门"].replace("张丽君", "药品1部", inplace=True)
   df1000["部门"].replace("王振杰", "药品1部", inplace=True)


   df1001=df1000.groupby(["部门"], as_index=False)[
        "年初余额", "1月应收余额", "2月应收余额", "3月应收余额","4月应收余额","5月应收余额","6月应收余额","7月应收余额","8月应收余额","9月应收余额",
        "10月应收余额","11月应收余额","12月应收余额"].sum();


   df1002 = pd.concat([df1000, df1001], ignore_index=True)  # 组合

   df1003 = df1002.fillna(0)
   df1003["业务员"].replace(0,"药品1部", inplace=True)

   print(df1003)

   df1004=df1003.groupby(["业务员"], as_index=False)[
       "年初余额", "1月应收余额", "2月应收余额", "3月应收余额", "4月应收余额", "5月应收余额", "6月应收余额", "7月应收余额", "8月应收余额", "9月应收余额",
       "10月应收余额", "11月应收余额", "12月应收余额"].sum();

   ####药品二部应收余额开始
   ####
   ####杨阳开始d
   dfd10 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   dfd11 = dfd10[dfd10["业务员"] == "杨阳"]
   dfd12 = dfd11.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方"], axis=1)  # 删列
   dfd13 = dfd12.rename(columns={'期末余额': '1月应收余额'});
   print(dfd13)
   dfd14 = df9[df9["期间"] == 2]
   dfd15 = dfd14[dfd14["业务员"] == "杨阳"]
   dfd16 = dfd15.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfd17 = dfd16.rename(columns={'期末余额': '2月应收余额'});
   print(dfd17)

   dfd18 = df9[df9["期间"] == 3]
   dfd19 = dfd18[dfd18["业务员"] == "杨阳"]
   dfd20 = dfd19.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfd21 = dfd20.rename(columns={'期末余额': '3月应收余额'});
   print(dfd21)

   dfd22 = df9[df9["期间"] == 4]
   dfd23 = dfd22[dfd22["业务员"] == "杨阳"]
   dfd24 = dfd23.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfd25 = dfd24.rename(columns={'期末余额': '4月应收余额'});
   print(dfd25)

   dfd26 = df9[df9["期间"] == 5]
   dfd27 = dfd26[dfa26["业务员"] == "杨阳"]
   dfd28 = dfd27.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfd29 = dfd28.rename(columns={'期末余额': '5月应收余额'});
   print(dfd29)

   dfd30 = df9[df9["期间"] == 6]
   dfd31 = dfd30[dfd30["业务员"] == "杨阳"]
   dfd32 = dfd31.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfd33 = dfd32.rename(columns={'期末余额': '6月应收余额'});
   print(dfd33)

   dfd34 = df9[df9["期间"] == 7]
   dfd35 = dfd34[dfd34["业务员"] == "杨阳"]
   dfd36 = dfd35.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfd37 = dfd36.rename(columns={'期末余额': '7月应收余额'});
   print(dfd37)

   dfd38 = df9[df9["期间"] == 8]
   dfd39 = dfd38[dfd38["业务员"] == "杨阳"]
   dfd40 = dfd39.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfd41 = dfd40.rename(columns={'期末余额': '8月应收余额'});
   print(dfd41)

   dfd42 = df9[df9["期间"] == 9]
   dfd43 = dfd42[dfd42["业务员"] == "杨阳"]
   dfd44 = dfd43.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfd45 = dfd44.rename(columns={'期末余额': '9月应收余额'});
   print(dfd45)

   dfd46 = df9[df9["期间"] == 10]
   dfd47 = dfd46[dfd46["业务员"] == "杨阳"]
   dfd48 = dfd47.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfd49 = dfd48.rename(columns={'期末余额': '10月应收余额'});
   print(dfd49)

   dfd50 = df9[df9["期间"] == 11]
   dfd51 = dfd50[dfd50["业务员"] == "杨阳"]
   dfd52 = dfd51.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfd53 = dfd52.rename(columns={'期末余额': '11月应收余额'});
   print(dfd53)

   dfd54 = df9[df9["期间"] == 12]
   dfd55 = dfd54[dfd54["业务员"] == "杨阳"]
   dfd56 = dfd55.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfd57 = dfd56.rename(columns={'期末余额': '12月应收余额'});
   print(dfd57)

   # df100 = pd.merge(df12, df15, right_index=True, left_index=True);
   dfd99 = pd.merge(dfd13, dfd17, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfd100 = pd.merge(dfd99, dfd21, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfd101 = pd.merge(dfd100, dfd25, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfd102 = pd.merge(dfd101, dfd29, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfd103 = pd.merge(dfd102, dfd33, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfd104 = pd.merge(dfd103, dfd37, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfd105 = pd.merge(dfd104, dfd41, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfd106 = pd.merge(dfd105, dfd45, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfd107 = pd.merge(dfd106, dfd49, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfd108 = pd.merge(dfd107, dfd53, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfd109 = pd.merge(dfd108, dfd57, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)

   print(dfd109)


   ####徐蕾开始e
   dfe10 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   dfe11 = dfe10[dfe10["业务员"] == "徐蕾"]
   dfe12 = dfe11.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方"], axis=1)  # 删列
   dfe13 = dfe12.rename(columns={'期末余额': '1月应收余额'});
   print(dfe13)
   dfe14 = df9[df9["期间"] == 2]
   dfe15 = dfe14[dfe14["业务员"] == "徐蕾"]
   dfe16 = dfe15.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfe17 = dfe16.rename(columns={'期末余额': '2月应收余额'});
   print(dfe17)

   dfe18 = df9[df9["期间"] == 3]
   dfe19 = dfe18[dfe18["业务员"] == "徐蕾"]
   dfe20 = dfe19.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfe21 = dfe20.rename(columns={'期末余额': '3月应收余额'});
   print(dfe21)

   dfe22 = df9[df9["期间"] == 4]
   dfe23 = dfe22[dfe22["业务员"] == "徐蕾"]
   dfe24 = dfe23.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfe25 = dfe24.rename(columns={'期末余额': '4月应收余额'});
   print(dfe25)

   dfe26 = df9[df9["期间"] == 5]
   dfe27 = dfe26[dfa26["业务员"] == "徐蕾"]
   dfe28 = dfe27.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfe29 = dfe28.rename(columns={'期末余额': '5月应收余额'});
   print(dfe29)

   dfe30 = df9[df9["期间"] == 6]
   dfe31 = dfe30[dfe30["业务员"] == "徐蕾"]
   dfe32 = dfe31.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfe33 = dfe32.rename(columns={'期末余额': '6月应收余额'});
   print(dfe33)

   dfe34 = df9[df9["期间"] == 7]
   dfe35 = dfe34[dfe34["业务员"] == "徐蕾"]
   dfe36 = dfe35.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfe37 = dfe36.rename(columns={'期末余额': '7月应收余额'});
   print(dfe37)

   dfe38 = df9[df9["期间"] == 8]
   dfe39 = dfe38[dfe38["业务员"] == "徐蕾"]
   dfe40 = dfe39.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfe41 = dfe40.rename(columns={'期末余额': '8月应收余额'});
   print(dfe41)

   dfe42 = df9[df9["期间"] == 9]
   dfe43 = dfe42[dfe42["业务员"] == "徐蕾"]
   dfe44 = dfe43.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfe45 = dfe44.rename(columns={'期末余额': '9月应收余额'});
   print(dfe45)

   dfe46 = df9[df9["期间"] == 10]
   dfe47 = dfe46[dfe46["业务员"] == "徐蕾"]
   dfe48 = dfe47.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfe49 = dfe48.rename(columns={'期末余额': '10月应收余额'});
   print(dfe49)

   dfe50 = df9[df9["期间"] == 11]
   dfe51 = dfe50[dfe50["业务员"] == "徐蕾"]
   dfe52 = dfe51.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfe53 = dfe52.rename(columns={'期末余额': '11月应收余额'});
   print(dfe53)

   dfe54 = df9[df9["期间"] == 12]
   dfe55 = dfe54[dfe54["业务员"] == "徐蕾"]
   dfe56 = dfe55.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfe57 = dfe56.rename(columns={'期末余额': '12月应收余额'});
   print(dfe57)

   # df100 = pd.merge(df12, df15, right_index=True, left_index=True);
   dfe99 = pd.merge(dfe13, dfe17, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfe100 = pd.merge(dfe99, dfe21, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfe101 = pd.merge(dfe100, dfe25, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfe102 = pd.merge(dfe101, dfe29, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfe103 = pd.merge(dfe102, dfe33, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfe104 = pd.merge(dfe103, dfe37, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfe105 = pd.merge(dfe104, dfe41, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfe106 = pd.merge(dfe105, dfe45, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfe107 = pd.merge(dfe106, dfe49, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfe108 = pd.merge(dfe107, dfe53, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfe109 = pd.merge(dfe108, dfe57, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)

   print(dfe109)

   ####王伟平开始f
   dff10 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   dff11 = dff10[dff10["业务员"] == "王伟平"]
   dff12 = dff11.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方"], axis=1)  # 删列
   dff13 = dff12.rename(columns={'期末余额': '1月应收余额'});
   print(dff13)
   dff14 = df9[df9["期间"] == 2]
   dff15 = dff14[dff14["业务员"] == "王伟平"]
   dff16 = dff15.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dff17 = dff16.rename(columns={'期末余额': '2月应收余额'});
   print(dff17)

   dff18 = df9[df9["期间"] == 3]
   dff19 = dff18[dff18["业务员"] == "王伟平"]
   dff20 = dff19.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dff21 = dff20.rename(columns={'期末余额': '3月应收余额'});
   print(dff21)

   dff22 = df9[df9["期间"] == 4]
   dff23 = dff22[dff22["业务员"] == "王伟平"]
   dff24 = dff23.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dff25 = dff24.rename(columns={'期末余额': '4月应收余额'});
   print(dff25)

   dff26 = df9[df9["期间"] == 5]
   dff27 = dff26[dfa26["业务员"] == "王伟平"]
   dff28 = dff27.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dff29 = dff28.rename(columns={'期末余额': '5月应收余额'});
   print(dff29)

   dff30 = df9[df9["期间"] == 6]
   dff31 = dff30[dff30["业务员"] == "王伟平"]
   dff32 = dff31.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dff33 = dff32.rename(columns={'期末余额': '6月应收余额'});
   print(dff33)

   dff34 = df9[df9["期间"] == 7]
   dff35 = dff34[dff34["业务员"] == "王伟平"]
   dff36 = dff35.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dff37 = dff36.rename(columns={'期末余额': '7月应收余额'});
   print(dff37)

   dff38 = df9[df9["期间"] == 8]
   dff39 = dff38[dff38["业务员"] == "王伟平"]
   dff40 = dff39.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dff41 = dff40.rename(columns={'期末余额': '8月应收余额'});
   print(dff41)

   dff42 = df9[df9["期间"] == 9]
   dff43 = dff42[dff42["业务员"] == "王伟平"]
   dff44 = dff43.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dff45 = dff44.rename(columns={'期末余额': '9月应收余额'});
   print(dff45)

   dff46 = df9[df9["期间"] == 10]
   dff47 = dff46[dff46["业务员"] == "王伟平"]
   dff48 = dff47.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dff49 = dff48.rename(columns={'期末余额': '10月应收余额'});
   print(dff49)

   dff50 = df9[df9["期间"] == 11]
   dff51 = dff50[dff50["业务员"] == "王伟平"]
   dff52 = dff51.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dff53 = dff52.rename(columns={'期末余额': '11月应收余额'});
   print(dff53)

   dff54 = df9[df9["期间"] == 12]
   dff55 = dff54[dff54["业务员"] == "王伟平"]
   dff56 = dff55.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dff57 = dff56.rename(columns={'期末余额': '12月应收余额'});
   print(dff57)

   # df100 = pd.merge(df12, df15, right_index=True, left_index=True);
   dff99 = pd.merge(dff13, dff17, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dff100 = pd.merge(dff99, dff21, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dff101 = pd.merge(dff100, dff25, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dff102 = pd.merge(dff101, dff29, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dff103 = pd.merge(dff102, dff33, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dff104 = pd.merge(dff103, dff37, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dff105 = pd.merge(dff104, dff41, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dff106 = pd.merge(dff105, dff45, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dff107 = pd.merge(dff106, dff49, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dff108 = pd.merge(dff107, dff53, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dff109 = pd.merge(dff108, dff57, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)

   print(dff109)

   ####王宇栋开始g
   dfg10 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   dfg11 = dfg10[dfg10["业务员"] == "王宇栋"]
   dfg12 = dfg11.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方"], axis=1)  # 删列
   dfg13 = dfg12.rename(columns={'期末余额': '1月应收余额'});
   print(dfg13)
   dfg14 = df9[df9["期间"] == 2]
   dfg15 = dfg14[dfg14["业务员"] == "王宇栋"]
   dfg16 = dfg15.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfg17 = dfg16.rename(columns={'期末余额': '2月应收余额'});
   print(dfg17)

   dfg18 = df9[df9["期间"] == 3]
   dfg19 = dfg18[dfg18["业务员"] == "王宇栋"]
   dfg20 = dfg19.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfg21 = dfg20.rename(columns={'期末余额': '3月应收余额'});
   print(dfg21)

   dfg22 = df9[df9["期间"] == 4]
   dfg23 = dfg22[dfg22["业务员"] == "王宇栋"]
   dfg24 = dfg23.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfg25 = dfg24.rename(columns={'期末余额': '4月应收余额'});
   print(dfg25)

   dfg26 = df9[df9["期间"] == 5]
   dfg27 = dfg26[dfa26["业务员"] == "王宇栋"]
   dfg28 = dfg27.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfg29 = dfg28.rename(columns={'期末余额': '5月应收余额'});
   print(dfg29)

   dfg30 = df9[df9["期间"] == 6]
   dfg31 = dfg30[dfg30["业务员"] == "王宇栋"]
   dfg32 = dfg31.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfg33 = dfg32.rename(columns={'期末余额': '6月应收余额'});
   print(dfg33)

   dfg34 = df9[df9["期间"] == 7]
   dfg35 = dfg34[dfg34["业务员"] == "王宇栋"]
   dfg36 = dfg35.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfg37 = dfg36.rename(columns={'期末余额': '7月应收余额'});
   print(dfg37)

   dfg38 = df9[df9["期间"] == 8]
   dfg39 = dfg38[dfg38["业务员"] == "王宇栋"]
   dfg40 = dfg39.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfg41 = dfg40.rename(columns={'期末余额': '8月应收余额'});
   print(dfg41)

   dfg42 = df9[df9["期间"] == 9]
   dfg43 = dfg42[dfg42["业务员"] == "王宇栋"]
   dfg44 = dfg43.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfg45 = dfg44.rename(columns={'期末余额': '9月应收余额'});
   print(dfg45)

   dfg46 = df9[df9["期间"] == 10]
   dfg47 = dfg46[dfg46["业务员"] == "王宇栋"]
   dfg48 = dfg47.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfg49 = dfg48.rename(columns={'期末余额': '10月应收余额'});
   print(dfg49)

   dfg50 = df9[df9["期间"] == 11]
   dfg51 = dfg50[dfg50["业务员"] == "王宇栋"]
   dfg52 = dfg51.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfg53 = dfg52.rename(columns={'期末余额': '11月应收余额'});
   print(dfg53)

   dfg54 = df9[df9["期间"] == 12]
   dfg55 = dfg54[dfg54["业务员"] == "王宇栋"]
   dfg56 = dfg55.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfg57 = dfg56.rename(columns={'期末余额': '12月应收余额'});
   print(dfg57)

   # df100 = pd.merge(df12, df15, right_index=True, left_index=True);
   dfg99 = pd.merge(dfg13, dfg17, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfg100 = pd.merge(dfg99, dfg21, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfg101 = pd.merge(dfg100, dfg25, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfg102 = pd.merge(dfg101, dfg29, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfg103 = pd.merge(dfg102, dfg33, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfg104 = pd.merge(dfg103, dfg37, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfg105 = pd.merge(dfg104, dfg41, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfg106 = pd.merge(dfg105, dfg45, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfg107 = pd.merge(dfg106, dfg49, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfg108 = pd.merge(dfg107, dfg53, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfg109 = pd.merge(dfg108, dfg57, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)

   print(dfg109)
   ####朱津齐开始g
   dfh10 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   dfh11 = dfh10[dfh10["业务员"] == "朱津齐"]
   dfh12 = dfh11.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方"], axis=1)  # 删列
   dfh13 = dfh12.rename(columns={'期末余额': '1月应收余额'});
   print(dfh13)
   dfh14 = df9[df9["期间"] == 2]
   dfh15 = dfh14[dfh14["业务员"] == "朱津齐"]
   dfh16 = dfh15.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfh17 = dfh16.rename(columns={'期末余额': '2月应收余额'});
   print(dfh17)

   dfh18 = df9[df9["期间"] == 3]
   dfh19 = dfh18[dfh18["业务员"] == "朱津齐"]
   dfh20 = dfh19.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfh21 = dfh20.rename(columns={'期末余额': '3月应收余额'});
   print(dfh21)

   dfh22 = df9[df9["期间"] == 4]
   dfh23 = dfh22[dfh22["业务员"] == "朱津齐"]
   dfh24 = dfh23.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfh25 = dfh24.rename(columns={'期末余额': '4月应收余额'});
   print(dfh25)

   dfh26 = df9[df9["期间"] == 5]
   dfh27 = dfh26[dfa26["业务员"] == "朱津齐"]
   dfh28 = dfh27.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfh29 = dfh28.rename(columns={'期末余额': '5月应收余额'});
   print(dfh29)

   dfh30 = df9[df9["期间"] == 6]
   dfh31 = dfh30[dfh30["业务员"] == "朱津齐"]
   dfh32 = dfh31.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfh33 = dfh32.rename(columns={'期末余额': '6月应收余额'});
   print(dfh33)

   dfh34 = df9[df9["期间"] == 7]
   dfh35 = dfh34[dfh34["业务员"] == "朱津齐"]
   dfh36 = dfh35.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfh37 = dfh36.rename(columns={'期末余额': '7月应收余额'});
   print(dfh37)

   dfh38 = df9[df9["期间"] == 8]
   dfh39 = dfh38[dfh38["业务员"] == "朱津齐"]
   dfh40 = dfh39.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfh41 = dfh40.rename(columns={'期末余额': '8月应收余额'});
   print(dfh41)

   dfh42 = df9[df9["期间"] == 9]
   dfh43 = dfh42[dfh42["业务员"] == "朱津齐"]
   dfh44 = dfh43.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfh45 = dfh44.rename(columns={'期末余额': '9月应收余额'});
   print(dfh45)

   dfh46 = df9[df9["期间"] == 10]
   dfh47 = dfh46[dfh46["业务员"] == "朱津齐"]
   dfh48 = dfh47.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfh49 = dfh48.rename(columns={'期末余额': '10月应收余额'});
   print(dfh49)

   dfh50 = df9[df9["期间"] == 11]
   dfh51 = dfh50[dfh50["业务员"] == "朱津齐"]
   dfh52 = dfh51.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfh53 = dfh52.rename(columns={'期末余额': '11月应收余额'});
   print(dfh53)

   dfh54 = df9[df9["期间"] == 12]
   dfh55 = dfh54[dfh54["业务员"] == "朱津齐"]
   dfh56 = dfh55.drop(["期间", "期初余额", "借方", "贷方", "本年累计借方", "本年累计贷方", "年初余额"], axis=1)  # 删列
   dfh57 = dfh56.rename(columns={'期末余额': '12月应收余额'});
   print(dfh57)

   # df100 = pd.merge(df12, df15, right_index=True, left_index=True);
   dfh99 = pd.merge(dfh13, dfh17, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfh100 = pd.merge(dfh99, dfh21, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfh101 = pd.merge(dfh100, dfh25, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfh102 = pd.merge(dfh101, dfh29, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfh103 = pd.merge(dfh102, dfh33, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfh104 = pd.merge(dfh103, dfh37, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfh105 = pd.merge(dfh104, dfh41, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfh106 = pd.merge(dfh105, dfh45, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfh107 = pd.merge(dfh106, dfh49, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfh108 = pd.merge(dfh107, dfh53, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)
   dfh109 = pd.merge(dfh108, dfh57, how='left', on=['业务员']);  # 完全相同合并，忽略没有的货品ID(没有how)

   print(dfh109)




   df1100 = pd.concat([dfd109, dfe109, dff109, dfg109,dfh109], ignore_index=True)  # 组合

   ####药品二部合计
   df1100["部门"] = df1100["业务员"]
   # df1001=df1000.rename(columns={'傅诗云': '药品1部','戚肆朝': '药品1部','张丽君': '药品1部','王振杰': '药品1部'});

   df1100["部门"].replace("杨阳", "药品2部", inplace=True)
   df1100["部门"].replace("徐蕾", "药品2部", inplace=True)
   df1100["部门"].replace("王伟平", "药品2部", inplace=True)
   df1100["部门"].replace("王宇栋", "药品2部", inplace=True)
   df1100["部门"].replace("朱津齐", "药品2部", inplace=True)

   df1101 = df1100.groupby(["部门"], as_index=False)[
       "年初余额", "1月应收余额", "2月应收余额", "3月应收余额", "4月应收余额", "5月应收余额", "6月应收余额", "7月应收余额", "8月应收余额", "9月应收余额",
       "10月应收余额", "11月应收余额", "12月应收余额"].sum();

   df1102 = pd.concat([df1100, df1101], ignore_index=True)  # 组合

   df1103 = df1102.fillna(0)
   df1103["业务员"].replace(0, "药品2部", inplace=True)

   print(df1103)

   df1104 = df1103.groupby(["业务员"], as_index=False)[
       "年初余额", "1月应收余额", "2月应收余额", "3月应收余额", "4月应收余额", "5月应收余额", "6月应收余额", "7月应收余额", "8月应收余额", "9月应收余额",
       "10月应收余额", "11月应收余额", "12月应收余额"].sum();



   df1200 =pd.concat([df1004, df1104], ignore_index=True)  # 组合
   print(df1200)

   #df1200.to_excel(excel_writer="d:/测试大神1206.xlsx",sheet_name="按发票汇总");


   df201a = df1200[df1200["业务员"] == "傅诗云"]
   df201a["1月平均余额"] = (df201a["1月应收余额"] + df201a["年初余额"]) / 2
   df201a["2月平均余额"] = (df201a["年初余额"] + df201a["2月应收余额"]) / 2
   df201a["3月平均余额"] = (df201a["年初余额"] + df201a["3月应收余额"]) / 2
   df201a["4月平均余额"] = (df201a["年初余额"] + df201a["4月应收余额"]) / 2
   df201a["5月平均余额"] = (df201a["年初余额"] + df201a["5月应收余额"]) / 2
   df201a["6月平均余额"] = (df201a["年初余额"] + df201a["6月应收余额"]) / 2
   df201a["7月平均余额"] = (df201a["年初余额"] + df201a["7月应收余额"]) / 2
   df201a["8月平均余额"] = (df201a["年初余额"] + df201a["8月应收余额"]) / 2
   df201a["9月平均余额"] = (df201a["年初余额"] + df201a["9月应收余额"]) / 2
   df201a["10月平均余额"] = (df201a["年初余额"] + df201a["10月应收余额"]) / 2
   df201a["11月平均余额"] = (df201a["年初余额"] + df201a["11月应收余额"]) / 2
   df201a["12月平均余额"] = (df201a["年初余额"] + df201a["12月应收余额"]) / 2
######戚肆朝df202a
   df202a = df1200[df1200["业务员"] == "戚肆朝"]
   df202a["1月平均余额"] = (df202a["1月应收余额"] + df202a["年初余额"]) / 2
   df202a["2月平均余额"] = (df202a["年初余额"] + df202a["2月应收余额"]) / 2
   df202a["3月平均余额"] = (df202a["年初余额"] + df202a["3月应收余额"]) / 2
   df202a["4月平均余额"] = (df202a["年初余额"] + df202a["4月应收余额"]) / 2
   df202a["5月平均余额"] = (df202a["年初余额"] + df202a["5月应收余额"]) / 2
   df202a["6月平均余额"] = (df202a["年初余额"] + df202a["6月应收余额"]) / 2
   df202a["7月平均余额"] = (df202a["年初余额"] + df202a["7月应收余额"]) / 2
   df202a["8月平均余额"] = (df202a["年初余额"] + df202a["8月应收余额"]) / 2
   df202a["9月平均余额"] = (df202a["年初余额"] + df202a["9月应收余额"]) / 2
   df202a["10月平均余额"] = (df202a["年初余额"] + df202a["10月应收余额"]) / 2
   df202a["11月平均余额"] = (df202a["年初余额"] + df202a["11月应收余额"]) / 2
   df202a["12月平均余额"] = (df202a["年初余额"] + df202a["12月应收余额"]) / 2
   #####张丽君
   df203a = df1200[df1200["业务员"] == "张丽君"]
   df203a["1月平均余额"] = (df203a["1月应收余额"] + df203a["年初余额"]) / 2
   df203a["2月平均余额"] = (df203a["年初余额"] + df203a["2月应收余额"]) / 2
   df203a["3月平均余额"] = (df203a["年初余额"] + df203a["3月应收余额"]) / 2
   df203a["4月平均余额"] = (df203a["年初余额"] + df203a["4月应收余额"]) / 2
   df203a["5月平均余额"] = (df203a["年初余额"] + df203a["5月应收余额"]) / 2
   df203a["6月平均余额"] = (df203a["年初余额"] + df203a["6月应收余额"]) / 2
   df203a["7月平均余额"] = (df203a["年初余额"] + df203a["7月应收余额"]) / 2
   df203a["8月平均余额"] = (df203a["年初余额"] + df203a["8月应收余额"]) / 2
   df203a["9月平均余额"] = (df203a["年初余额"] + df203a["9月应收余额"]) / 2
   df203a["10月平均余额"] = (df203a["年初余额"] + df203a["10月应收余额"]) / 2
   df203a["11月平均余额"] = (df203a["年初余额"] + df203a["11月应收余额"]) / 2
   df203a["12月平均余额"] = (df203a["年初余额"] + df203a["12月应收余额"]) / 2


#####王振杰
   df205a = df1200[df1200["业务员"] == "王振杰"]
   df205a["1月平均余额"] = (df205a["1月应收余额"] + df205a["年初余额"]) / 2
   df205a["2月平均余额"] = (df205a["年初余额"] + df205a["2月应收余额"]) / 2
   df205a["3月平均余额"] = (df205a["年初余额"] + df205a["3月应收余额"]) / 2
   df205a["4月平均余额"] = (df205a["年初余额"] + df205a["4月应收余额"]) / 2
   df205a["5月平均余额"] = (df205a["年初余额"] + df205a["5月应收余额"]) / 2
   df205a["6月平均余额"] = (df205a["年初余额"] + df205a["6月应收余额"]) / 2
   df205a["7月平均余额"] = (df205a["年初余额"] + df205a["7月应收余额"]) / 2
   df205a["8月平均余额"] = (df205a["年初余额"] + df205a["8月应收余额"]) / 2
   df205a["9月平均余额"] = (df205a["年初余额"] + df205a["9月应收余额"]) / 2
   df205a["10月平均余额"] = (df205a["年初余额"] + df205a["10月应收余额"]) / 2
   df205a["11月平均余额"] = (df205a["年初余额"] + df205a["11月应收余额"]) / 2
   df205a["12月平均余额"] = (df205a["年初余额"] + df205a["12月应收余额"]) / 2
####药品1部平均余额小计
   df206a = df1200[df1200["业务员"] == "药品1部"]
   df206a["1月平均余额"] = (df206a["1月应收余额"] + df206a["年初余额"]) / 2
   df206a["2月平均余额"] = (df206a["年初余额"] + df206a["2月应收余额"]) / 2
   df206a["3月平均余额"] = (df206a["年初余额"] + df206a["3月应收余额"]) / 2
   df206a["4月平均余额"] = (df206a["年初余额"] + df206a["4月应收余额"]) / 2
   df206a["5月平均余额"] = (df206a["年初余额"] + df206a["5月应收余额"]) / 2
   df206a["6月平均余额"] = (df206a["年初余额"] + df206a["6月应收余额"]) / 2
   df206a["7月平均余额"] = (df206a["年初余额"] + df206a["7月应收余额"]) / 2
   df206a["8月平均余额"] = (df206a["年初余额"] + df206a["8月应收余额"]) / 2
   df206a["9月平均余额"] = (df206a["年初余额"] + df206a["9月应收余额"]) / 2
   df206a["10月平均余额"] = (df206a["年初余额"] + df206a["10月应收余额"]) / 2
   df206a["11月平均余额"] = (df206a["年初余额"] + df206a["11月应收余额"]) / 2
   df206a["12月平均余额"] = (df206a["年初余额"] + df206a["12月应收余额"]) / 2
#####徐蕾
   df207a = df1200[df1200["业务员"] == "徐蕾"]
   df207a["1月平均余额"] = (df207a["1月应收余额"] + df207a["年初余额"]) / 2
   df207a["2月平均余额"] = (df207a["年初余额"] + df207a["2月应收余额"]) / 2
   df207a["3月平均余额"] = (df207a["年初余额"] + df207a["3月应收余额"]) / 2
   df207a["4月平均余额"] = (df207a["年初余额"] + df207a["4月应收余额"]) / 2
   df207a["5月平均余额"] = (df207a["年初余额"] + df207a["5月应收余额"]) / 2
   df207a["6月平均余额"] = (df207a["年初余额"] + df207a["6月应收余额"]) / 2
   df207a["7月平均余额"] = (df207a["年初余额"] + df207a["7月应收余额"]) / 2
   df207a["8月平均余额"] = (df207a["年初余额"] + df207a["8月应收余额"]) / 2
   df207a["9月平均余额"] = (df207a["年初余额"] + df207a["9月应收余额"]) / 2
   df207a["10月平均余额"] = (df207a["年初余额"] + df207a["10月应收余额"]) / 2
   df207a["11月平均余额"] = (df207a["年初余额"] + df207a["11月应收余额"]) / 2
   df207a["12月平均余额"] = (df207a["年初余额"] + df207a["12月应收余额"]) / 2

#####朱津齐
   df208a = df1200[df1200["业务员"] == "朱津齐"]
   df208a["1月平均余额"] = (df208a["1月应收余额"] + df208a["年初余额"]) / 2
   df208a["2月平均余额"] = (df208a["年初余额"] + df208a["2月应收余额"]) / 2
   df208a["3月平均余额"] = (df208a["年初余额"] + df208a["3月应收余额"]) / 2
   df208a["4月平均余额"] = (df208a["年初余额"] + df208a["4月应收余额"]) / 2
   df208a["5月平均余额"] = (df208a["年初余额"] + df208a["5月应收余额"]) / 2
   df208a["6月平均余额"] = (df208a["年初余额"] + df208a["6月应收余额"]) / 2
   df208a["7月平均余额"] = (df208a["年初余额"] + df208a["7月应收余额"]) / 2
   df208a["8月平均余额"] = (df208a["年初余额"] + df208a["8月应收余额"]) / 2
   df208a["9月平均余额"] = (df208a["年初余额"] + df208a["9月应收余额"]) / 2
   df208a["10月平均余额"] = (df208a["年初余额"] + df208a["10月应收余额"]) / 2
   df208a["11月平均余额"] = (df208a["年初余额"] + df208a["11月应收余额"]) / 2
   df208a["12月平均余额"] = (df208a["年初余额"] + df208a["12月应收余额"]) / 2

####杨阳
   df209a = df1200[df1200["业务员"] == "杨阳"]
   df209a["1月平均余额"] = (df209a["1月应收余额"] + df209a["年初余额"]) / 2
   df209a["2月平均余额"] = (df209a["年初余额"] + df209a["2月应收余额"]) / 2
   df209a["3月平均余额"] = (df209a["年初余额"] + df209a["3月应收余额"]) / 2
   df209a["4月平均余额"] = (df209a["年初余额"] + df209a["4月应收余额"]) / 2
   df209a["5月平均余额"] = (df209a["年初余额"] + df209a["5月应收余额"]) / 2
   df209a["6月平均余额"] = (df209a["年初余额"] + df209a["6月应收余额"]) / 2
   df209a["7月平均余额"] = (df209a["年初余额"] + df209a["7月应收余额"]) / 2
   df209a["8月平均余额"] = (df209a["年初余额"] + df209a["8月应收余额"]) / 2
   df209a["9月平均余额"] = (df209a["年初余额"] + df209a["9月应收余额"]) / 2
   df209a["10月平均余额"] = (df209a["年初余额"] + df209a["10月应收余额"]) / 2
   df209a["11月平均余额"] = (df209a["年初余额"] + df209a["11月应收余额"]) / 2
   df209a["12月平均余额"] = (df209a["年初余额"] + df209a["12月应收余额"]) / 2
####王伟平
   df210a = df1200[df1200["业务员"] == "王伟平"]
   df210a["1月平均余额"] = (df210a["1月应收余额"] + df210a["年初余额"]) / 2
   df210a["2月平均余额"] = (df210a["年初余额"] + df210a["2月应收余额"]) / 2
   df210a["3月平均余额"] = (df210a["年初余额"] + df210a["3月应收余额"]) / 2
   df210a["4月平均余额"] = (df210a["年初余额"] + df210a["4月应收余额"]) / 2
   df210a["5月平均余额"] = (df210a["年初余额"] + df210a["5月应收余额"]) / 2
   df210a["6月平均余额"] = (df210a["年初余额"] + df210a["6月应收余额"]) / 2
   df210a["7月平均余额"] = (df210a["年初余额"] + df210a["7月应收余额"]) / 2
   df210a["8月平均余额"] = (df210a["年初余额"] + df210a["8月应收余额"]) / 2
   df210a["9月平均余额"] = (df210a["年初余额"] + df210a["9月应收余额"]) / 2
   df210a["10月平均余额"] = (df210a["年初余额"] + df210a["10月应收余额"]) / 2
   df210a["11月平均余额"] = (df210a["年初余额"] + df210a["11月应收余额"]) / 2
   df210a["12月平均余额"] = (df210a["年初余额"] + df210a["12月应收余额"]) / 2
####王宇栋
   df211a = df1200[df1200["业务员"] == "王宇栋"]
   df211a["1月平均余额"] = (df211a["1月应收余额"] + df211a["年初余额"]) / 2
   df211a["2月平均余额"] = (df211a["年初余额"] + df211a["2月应收余额"]) / 2
   df211a["3月平均余额"] = (df211a["年初余额"] + df211a["3月应收余额"]) / 2
   df211a["4月平均余额"] = (df211a["年初余额"] + df211a["4月应收余额"]) / 2
   df211a["5月平均余额"] = (df211a["年初余额"] + df211a["5月应收余额"]) / 2
   df211a["6月平均余额"] = (df211a["年初余额"] + df211a["6月应收余额"]) / 2
   df211a["7月平均余额"] = (df211a["年初余额"] + df211a["7月应收余额"]) / 2
   df211a["8月平均余额"] = (df211a["年初余额"] + df211a["8月应收余额"]) / 2
   df211a["9月平均余额"] = (df211a["年初余额"] + df211a["9月应收余额"]) / 2
   df211a["10月平均余额"] = (df211a["年初余额"] + df211a["10月应收余额"]) / 2
   df211a["11月平均余额"] = (df211a["年初余额"] + df211a["11月应收余额"]) / 2
   df211a["12月平均余额"] = (df211a["年初余额"] + df211a["12月应收余额"]) / 2

#####药品2部
   df212a = df1200[df1200["业务员"] == "药品2部"]
   df212a["1月平均余额"] = (df212a["1月应收余额"] + df212a["年初余额"]) / 2
   df212a["2月平均余额"] = (df212a["年初余额"] + df212a["2月应收余额"]) / 2
   df212a["3月平均余额"] = (df212a["年初余额"] + df212a["3月应收余额"]) / 2
   df212a["4月平均余额"] = (df212a["年初余额"] + df212a["4月应收余额"]) / 2
   df212a["5月平均余额"] = (df212a["年初余额"] + df212a["5月应收余额"]) / 2
   df212a["6月平均余额"] = (df212a["年初余额"] + df212a["6月应收余额"]) / 2
   df212a["7月平均余额"] = (df212a["年初余额"] + df212a["7月应收余额"]) / 2
   df212a["8月平均余额"] = (df212a["年初余额"] + df212a["8月应收余额"]) / 2
   df212a["9月平均余额"] = (df212a["年初余额"] + df212a["9月应收余额"]) / 2
   df212a["10月平均余额"] = (df212a["年初余额"] + df212a["10月应收余额"]) / 2
   df212a["11月平均余额"] = (df212a["年初余额"] + df212a["11月应收余额"]) / 2
   df212a["12月平均余额"] = (df212a["年初余额"] + df212a["12月应收余额"]) / 2

   #####药品合计收款天数

   #df213a = df1200[df1200["业务员"] != "傅诗云"]
   #df214a = df213a[df213a["业务员"] != "张丽君"]
   #df215a = df214a[df214a["业务员"] != "戚肆朝"]
   #df216a = df215a[df215a["业务员"] != "王振杰"]
   #df217a = df216a[df216a["业务员"] != "徐蕾"]
   #df218a = df217a[df217a["业务员"] != "朱津齐"]
   #df219a = df218a[df218a["业务员"] != "杨阳"]
   #df220a = df219a[df219a["业务员"] != "王伟平"]
   #df221a = df220a[df220a["业务员"] != "王宇栋"]
   #df213a = df1200[df1200["业务员"] == "药品2部"]




   #df213a["业务员"].replace("药品2部", "合计", inplace=True)
   df1299=pd.concat([df201a, df202a, df203a,df205a,df206a,df207a,df208a,df209a,df210a,df211a,df212a], ignore_index=True)  # 组合
   df1300=df1299.drop(["年初余额","1月应收余额", "2月应收余额","3月应收余额","4月应收余额","5月应收余额","6月应收余额","7月应收余额","8月应收余额","9月应收余额","10月应收余额","11月应收余额","12月应收余额"], axis=1)  # 删列

   #df213a.to_excel(excel_writer="d:/测试大神1207.xlsx", sheet_name="按发票汇总");





#####表3销售额开始
   #####傅诗云
   df10aa = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   df11aa = df10aa[df10aa["业务员"] == "傅诗云"]
   df12aa = df11aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方","期末余额","年初余额"], axis=1)  # 删列
   df13aa = df12aa.rename(columns={'借方': '1月销售'});
   print(df13aa)
   df14aa = df9[df9["期间"] == 2]  # df63[df63["部门"] == "11部"]
   df15aa = df14aa[df14aa["业务员"] == "傅诗云"]
   df16aa = df15aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额","年初余额"], axis=1)  # 删列
   df17aa = df16aa.rename(columns={'借方': '2月销售'});
   print(df17aa)
   df18aa = df9[df9["期间"] == 3]  # df63[df63["部门"] == "11部"]
   df19aa = df18aa[df18aa["业务员"] == "傅诗云"]
   df20aa = df19aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df21aa = df20aa.rename(columns={'借方': '3月销售'});
   print(df21aa)
   df22aa = df9[df9["期间"] == 4]  # df63[df63["部门"] == "11部"]
   df23aa = df22aa[df22aa["业务员"] == "傅诗云"]
   df24aa = df23aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df25aa = df24aa.rename(columns={'借方': '4月销售'});
   print(df25aa)
   df26aa = df9[df9["期间"] == 5]  # df63[df63["部门"] == "11部"]
   df27aa = df26aa[df26aa["业务员"] == "傅诗云"]
   df28aa = df27aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df29aa = df28aa.rename(columns={'借方': '5月销售'});
   print(df29aa)
   df30aa = df9[df9["期间"] == 6]  # df63[df63["部门"] == "11部"]
   df31aa = df30aa[df30aa["业务员"] == "傅诗云"]
   df32aa = df31aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df33aa = df32aa.rename(columns={'借方': '6月销售'});
   print(df33aa)
   df34aa = df9[df9["期间"] == 7]  # df63[df63["部门"] == "11部"]
   df35aa = df34aa[df34aa["业务员"] == "傅诗云"]
   df36aa = df35aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df37aa = df36aa.rename(columns={'借方': '7月销售'});
   print(df37aa)
   df38aa = df9[df9["期间"] == 8]  # df63[df63["部门"] == "11部"]
   df39aa = df38aa[df38aa["业务员"] == "傅诗云"]
   df40aa = df39aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df41aa = df40aa.rename(columns={'借方': '8月销售'});
   print(df41aa)
   df42aa = df9[df9["期间"] == 9]  # df63[df63["部门"] == "11部"]
   df43aa = df42aa[df42aa["业务员"] == "傅诗云"]
   df44aa = df43aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df45aa = df44aa.rename(columns={'借方': '9月销售'});
   print(df45aa)
   df46aa = df9[df9["期间"] == 10]  # df63[df63["部门"] == "11部"]
   df47aa = df46aa[df46aa["业务员"] == "傅诗云"]
   df48aa = df47aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df49aa = df48aa.rename(columns={'借方': '10月销售'});
   print(df49aa)
   df50aa = df9[df9["期间"] == 11]  # df63[df63["部门"] == "11部"]
   df51aa = df50aa[df50aa["业务员"] == "傅诗云"]
   df52aa = df51aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df53aa = df52aa.rename(columns={'借方': '11月销售'});
   print(df53aa)
   df54aa = df9[df9["期间"] == 12]  # df63[df63["部门"] == "11部"]
   df55aa = df54aa[df54aa["业务员"] == "傅诗云"]
   df56aa = df55aa.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df57aa = df56aa.rename(columns={'借方': '12月销售'});
   print(df57aa)

   df100aa=pd.merge(df13aa, df17aa, how='left', on=['业务员']);#1+2
   df101aa = pd.merge(df100aa, df21aa, how='left', on=['业务员']);  # +3
   df102aa = pd.merge(df101aa, df25aa, how='left', on=['业务员']);  # +4
   df103aa = pd.merge(df102aa, df29aa, how='left', on=['业务员']);  # +5
   df104aa = pd.merge(df103aa, df33aa, how='left', on=['业务员']);  # +6
   df105aa = pd.merge(df104aa, df37aa, how='left', on=['业务员']);  # +7
   df106aa = pd.merge(df105aa, df41aa, how='left', on=['业务员']);  # +8
   df107aa = pd.merge(df106aa, df45aa, how='left', on=['业务员']);  # +9
   df108aa = pd.merge(df107aa, df49aa, how='left', on=['业务员']);  # +10
   df109aa = pd.merge(df108aa, df53aa, how='left', on=['业务员']);  # +11
   df110aa = pd.merge(df109aa, df57aa, how='left', on=['业务员']);  # +12

#####戚肆朝
   df10a1 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   df11a1 = df10a1[df10a1["业务员"] == "戚肆朝"]
   df12a1 = df11a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df13a1 = df12a1.rename(columns={'借方': '1月销售'});
   print(df13a1)
   df14a1 = df9[df9["期间"] == 2]  # df63[df63["部门"] == "11部"]
   df15a1 = df14a1[df14a1["业务员"] == "戚肆朝"]
   df16a1 = df15a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df17a1 = df16a1.rename(columns={'借方': '2月销售'});
   print(df17a1)
   df18a1 = df9[df9["期间"] == 3]  # df63[df63["部门"] == "11部"]
   df19a1 = df18a1[df18a1["业务员"] == "戚肆朝"]
   df20a1 = df19a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df21a1 = df20a1.rename(columns={'借方': '3月销售'});
   print(df21a1)
   df22a1 = df9[df9["期间"] == 4]  # df63[df63["部门"] == "11部"]
   df23a1 = df22a1[df22a1["业务员"] == "戚肆朝"]
   df24a1 = df23a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df25a1 = df24a1.rename(columns={'借方': '4月销售'});
   print(df25a1)
   df26a1 = df9[df9["期间"] == 5]  # df63[df63["部门"] == "11部"]
   df27a1 = df26a1[df26a1["业务员"] == "戚肆朝"]
   df28a1 = df27a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df29a1 = df28a1.rename(columns={'借方': '5月销售'});
   print(df29a1)
   df30a1 = df9[df9["期间"] == 6]  # df63[df63["部门"] == "11部"]
   df31a1 = df30a1[df30a1["业务员"] == "戚肆朝"]
   df32a1 = df31a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df33a1 = df32a1.rename(columns={'借方': '6月销售'});
   print(df33a1)
   df34a1 = df9[df9["期间"] == 7]  # df63[df63["部门"] == "11部"]
   df35a1 = df34a1[df34a1["业务员"] == "戚肆朝"]
   df36a1 = df35a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df37a1 = df36a1.rename(columns={'借方': '7月销售'});
   print(df37a1)
   df38a1 = df9[df9["期间"] == 8]  # df63[df63["部门"] == "11部"]
   df39a1 = df38a1[df38a1["业务员"] == "戚肆朝"]
   df40a1 = df39a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df41a1 = df40a1.rename(columns={'借方': '8月销售'});
   print(df41a1)
   df42a1 = df9[df9["期间"] == 9]  # df63[df63["部门"] == "11部"]
   df43a1 = df42a1[df42a1["业务员"] == "戚肆朝"]
   df44a1 = df43a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df45a1 = df44a1.rename(columns={'借方': '9月销售'});
   print(df45a1)
   df46a1 = df9[df9["期间"] == 10]  # df63[df63["部门"] == "11部"]
   df47a1 = df46a1[df46a1["业务员"] == "戚肆朝"]
   df48a1 = df47a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df49a1 = df48a1.rename(columns={'借方': '10月销售'});
   print(df49a1)
   df50a1 = df9[df9["期间"] == 11]  # df63[df63["部门"] == "11部"]
   df51a1 = df50a1[df50a1["业务员"] == "戚肆朝"]
   df52a1 = df51a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df53a1 = df52a1.rename(columns={'借方': '11月销售'});
   print(df53a1)
   df54a1 = df9[df9["期间"] == 12]  # df63[df63["部门"] == "11部"]
   df55a1 = df54a1[df54a1["业务员"] == "戚肆朝"]
   df56a1 = df55a1.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df57a1 = df56a1.rename(columns={'借方': '12月销售'});
   print(df57a1)

   df100a1 = pd.merge(df13a1, df17a1, how='left', on=['业务员']);  # 1+2
   df101a1 = pd.merge(df100a1, df21a1, how='left', on=['业务员']);  # +3
   df102a1 = pd.merge(df101a1, df25a1, how='left', on=['业务员']);  # +4
   df103a1 = pd.merge(df102a1, df29a1, how='left', on=['业务员']);  # +5
   df104a1 = pd.merge(df103a1, df33a1, how='left', on=['业务员']);  # +6
   df105a1 = pd.merge(df104a1, df37a1, how='left', on=['业务员']);  # +7
   df106a1 = pd.merge(df105a1, df41a1, how='left', on=['业务员']);  # +8
   df107a1 = pd.merge(df106a1, df45a1, how='left', on=['业务员']);  # +9
   df108a1 = pd.merge(df107a1, df49a1, how='left', on=['业务员']);  # +10
   df109a1 = pd.merge(df108a1, df53a1, how='left', on=['业务员']);  # +11
   df110a1 = pd.merge(df109a1, df57a1, how='left', on=['业务员']);  # +12


#####张丽君
   df10a2 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   df11a2 = df10a2[df10a2["业务员"] == "张丽君"]
   df12a2 = df11a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df13a2 = df12a2.rename(columns={'借方': '1月销售'});
   print(df13a2)
   df14a2 = df9[df9["期间"] == 2]  # df63[df63["部门"] == "11部"]
   df15a2 = df14a2[df14a2["业务员"] == "张丽君"]
   df16a2 = df15a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df17a2 = df16a2.rename(columns={'借方': '2月销售'});
   print(df17a2)
   df18a2 = df9[df9["期间"] == 3]  # df63[df63["部门"] == "11部"]
   df19a2 = df18a2[df18a2["业务员"] == "张丽君"]
   df20a2 = df19a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df21a2 = df20a2.rename(columns={'借方': '3月销售'});
   print(df21a2)
   df22a2 = df9[df9["期间"] == 4]  # df63[df63["部门"] == "11部"]
   df23a2 = df22a2[df22a2["业务员"] == "张丽君"]
   df24a2 = df23a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df25a2 = df24a2.rename(columns={'借方': '4月销售'});
   print(df25a2)
   df26a2 = df9[df9["期间"] == 5]  # df63[df63["部门"] == "11部"]
   df27a2 = df26a2[df26a2["业务员"] == "张丽君"]
   df28a2 = df27a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df29a2 = df28a2.rename(columns={'借方': '5月销售'});
   print(df29a2)
   df30a2 = df9[df9["期间"] == 6]  # df63[df63["部门"] == "11部"]
   df31a2 = df30a2[df30a2["业务员"] == "张丽君"]
   df32a2 = df31a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df33a2 = df32a2.rename(columns={'借方': '6月销售'});
   print(df33a2)
   df34a2 = df9[df9["期间"] == 7]  # df63[df63["部门"] == "11部"]
   df35a2 = df34a2[df34a2["业务员"] == "张丽君"]
   df36a2 = df35a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df37a2 = df36a2.rename(columns={'借方': '7月销售'});
   print(df37a2)
   df38a2 = df9[df9["期间"] == 8]  # df63[df63["部门"] == "11部"]
   df39a2 = df38a2[df38a2["业务员"] == "张丽君"]
   df40a2 = df39a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df41a2 = df40a2.rename(columns={'借方': '8月销售'});
   print(df41a2)
   df42a2 = df9[df9["期间"] == 9]  # df63[df63["部门"] == "11部"]
   df43a2 = df42a2[df42a2["业务员"] == "张丽君"]
   df44a2 = df43a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df45a2 = df44a2.rename(columns={'借方': '9月销售'});
   print(df45a2)
   df46a2 = df9[df9["期间"] == 10]  # df63[df63["部门"] == "11部"]
   df47a2 = df46a2[df46a2["业务员"] == "张丽君"]
   df48a2 = df47a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df49a2 = df48a2.rename(columns={'借方': '10月销售'});
   print(df49a2)
   df50a2 = df9[df9["期间"] == 11]  # df63[df63["部门"] == "11部"]
   df51a2 = df50a2[df50a2["业务员"] == "张丽君"]
   df52a2 = df51a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df53a2 = df52a2.rename(columns={'借方': '11月销售'});
   print(df53a2)
   df54a2 = df9[df9["期间"] == 12]  # df63[df63["部门"] == "11部"]
   df55a2 = df54a2[df54a2["业务员"] == "张丽君"]
   df56a2 = df55a2.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df57a2 = df56a2.rename(columns={'借方': '12月销售'});
   print(df57a2)

   df100a2 = pd.merge(df13a2, df17a2, how='left', on=['业务员']);  # 1+2
   df101a2 = pd.merge(df100a2, df21a2, how='left', on=['业务员']);  # +3
   df102a2 = pd.merge(df101a2, df25a2, how='left', on=['业务员']);  # +4
   df103a2 = pd.merge(df102a2, df29a2, how='left', on=['业务员']);  # +5
   df104a2 = pd.merge(df103a2, df33a2, how='left', on=['业务员']);  # +6
   df105a2 = pd.merge(df104a2, df37a2, how='left', on=['业务员']);  # +7
   df106a2 = pd.merge(df105a2, df41a2, how='left', on=['业务员']);  # +8
   df107a2 = pd.merge(df106a2, df45a2, how='left', on=['业务员']);  # +9
   df108a2 = pd.merge(df107a2, df49a2, how='left', on=['业务员']);  # +10
   df109a2 = pd.merge(df108a2, df53a2, how='left', on=['业务员']);  # +11
   df110a2 = pd.merge(df109a2, df57a2, how='left', on=['业务员']);  # +12
######王振杰
   df10a3 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   df11a3 = df10a3[df10a3["业务员"] == "王振杰"]
   df12a3 = df11a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df13a3 = df12a3.rename(columns={'借方': '1月销售'});
   print(df13a3)
   df14a3 = df9[df9["期间"] == 2]  # df63[df63["部门"] == "11部"]
   df15a3 = df14a3[df14a3["业务员"] == "王振杰"]
   df16a3 = df15a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df17a3 = df16a3.rename(columns={'借方': '2月销售'});
   print(df17a3)
   df18a3 = df9[df9["期间"] == 3]  # df63[df63["部门"] == "11部"]
   df19a3 = df18a3[df18a3["业务员"] == "王振杰"]
   df20a3 = df19a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df21a3 = df20a3.rename(columns={'借方': '3月销售'});
   print(df21a3)
   df22a3 = df9[df9["期间"] == 4]  # df63[df63["部门"] == "11部"]
   df23a3 = df22a3[df22a3["业务员"] == "王振杰"]
   df24a3 = df23a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df25a3 = df24a3.rename(columns={'借方': '4月销售'});
   print(df25a3)
   df26a3 = df9[df9["期间"] == 5]  # df63[df63["部门"] == "11部"]
   df27a3 = df26a3[df26a3["业务员"] == "王振杰"]
   df28a3 = df27a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df29a3 = df28a3.rename(columns={'借方': '5月销售'});
   print(df29a3)
   df30a3 = df9[df9["期间"] == 6]  # df63[df63["部门"] == "11部"]
   df31a3 = df30a3[df30a3["业务员"] == "王振杰"]
   df32a3 = df31a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df33a3 = df32a3.rename(columns={'借方': '6月销售'});
   print(df33a3)
   df34a3 = df9[df9["期间"] == 7]  # df63[df63["部门"] == "11部"]
   df35a3 = df34a3[df34a3["业务员"] == "王振杰"]
   df36a3 = df35a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df37a3 = df36a3.rename(columns={'借方': '7月销售'});
   print(df37a3)
   df38a3 = df9[df9["期间"] == 8]  # df63[df63["部门"] == "11部"]
   df39a3 = df38a3[df38a3["业务员"] == "王振杰"]
   df40a3 = df39a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df41a3 = df40a3.rename(columns={'借方': '8月销售'});
   print(df41a3)
   df42a3 = df9[df9["期间"] == 9]  # df63[df63["部门"] == "11部"]
   df43a3 = df42a3[df42a3["业务员"] == "王振杰"]
   df44a3 = df43a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df45a3 = df44a3.rename(columns={'借方': '9月销售'});
   print(df45a3)
   df46a3 = df9[df9["期间"] == 10]  # df63[df63["部门"] == "11部"]
   df47a3 = df46a3[df46a3["业务员"] == "王振杰"]
   df48a3 = df47a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df49a3 = df48a3.rename(columns={'借方': '10月销售'});
   print(df49a3)
   df50a3 = df9[df9["期间"] == 11]  # df63[df63["部门"] == "11部"]
   df51a3 = df50a3[df50a3["业务员"] == "王振杰"]
   df52a3 = df51a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df53a3 = df52a3.rename(columns={'借方': '11月销售'});
   print(df53a3)
   df54a3 = df9[df9["期间"] == 12]  # df63[df63["部门"] == "11部"]
   df55a3 = df54a3[df54a3["业务员"] == "王振杰"]
   df56a3 = df55a3.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df57a3 = df56a3.rename(columns={'借方': '12月销售'});
   print(df57a3)

   df100a3 = pd.merge(df13a3, df17a3, how='left', on=['业务员']);  # 1+2
   df101a3 = pd.merge(df100a3, df21a3, how='left', on=['业务员']);  # +3
   df102a3 = pd.merge(df101a3, df25a3, how='left', on=['业务员']);  # +4
   df103a3 = pd.merge(df102a3, df29a3, how='left', on=['业务员']);  # +5
   df104a3 = pd.merge(df103a3, df33a3, how='left', on=['业务员']);  # +6
   df105a3 = pd.merge(df104a3, df37a3, how='left', on=['业务员']);  # +7
   df106a3 = pd.merge(df105a3, df41a3, how='left', on=['业务员']);  # +8
   df107a3 = pd.merge(df106a3, df45a3, how='left', on=['业务员']);  # +9
   df108a3 = pd.merge(df107a3, df49a3, how='left', on=['业务员']);  # +10
   df109a3 = pd.merge(df108a3, df53a3, how='left', on=['业务员']);  # +11
   df110a3 = pd.merge(df109a3, df57a3, how='left', on=['业务员']);  # +12

   #####药品一部销售小计
   df1100a=pd.concat([df110aa, df110a1, df110a2, df110a3], ignore_index=True)  # 组合

   df1100a["部门"] = df1100a["业务员"]
   # df1001=df1000.rename(columns={'傅诗云': '药品1部','戚肆朝': '药品1部','张丽君': '药品1部','王振杰': '药品1部'});

   df1100a["部门"].replace("傅诗云", "药品1部", inplace=True)
   df1100a["部门"].replace("戚肆朝", "药品1部", inplace=True)
   df1100a["部门"].replace("张丽君", "药品1部", inplace=True)
   df1100a["部门"].replace("王振杰", "药品1部", inplace=True)
   #df1100a["部门"].replace("朱津齐", "药品2部", inplace=True)

   df1101a = df1100a.groupby(["部门"], as_index=False)[
       "1月销售", "2月销售", "3月销售", "4月销售", "5月销售", "6月销售", "7月销售", "8月销售", "9月销售", "10月销售",
       "11月销售", "12月销售"].sum();

   df1102a = pd.concat([df1100a, df1101a], ignore_index=True)  # 组合

   df1103a = df1102a.fillna(0)
   df1103a["业务员"].replace(0, "药品1部", inplace=True)

   print(df1103a)

   df1104a = df1103a.groupby(["业务员"], as_index=False)[
       "1月销售", "2月销售", "3月销售", "4月销售", "5月销售", "6月销售", "7月销售", "8月销售", "9月销售", "10月销售",
       "11月销售", "12月销售"].sum();

   #df1400 = pd.concat([df1004, df1104], ignore_index=True)  # 组合

#######杨阳
   df10a4 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   df11a4 = df10a4[df10a4["业务员"] == "杨阳"]
   df12a4 = df11a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df13a4 = df12a4.rename(columns={'借方': '1月销售'});
   print(df13a4)
   df14a4 = df9[df9["期间"] == 2]  # df63[df63["部门"] == "11部"]
   df15a4 = df14a4[df14a4["业务员"] == "杨阳"]
   df16a4 = df15a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df17a4 = df16a4.rename(columns={'借方': '2月销售'});
   print(df17a4)
   df18a4 = df9[df9["期间"] == 3]  # df63[df63["部门"] == "11部"]
   df19a4 = df18a4[df18a4["业务员"] == "杨阳"]
   df20a4 = df19a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df21a4 = df20a4.rename(columns={'借方': '3月销售'});
   print(df21a4)
   df22a4 = df9[df9["期间"] == 4]  # df63[df63["部门"] == "11部"]
   df23a4 = df22a4[df22a4["业务员"] == "杨阳"]
   df24a4 = df23a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df25a4 = df24a4.rename(columns={'借方': '4月销售'});
   print(df25a4)
   df26a4 = df9[df9["期间"] == 5]  # df63[df63["部门"] == "11部"]
   df27a4 = df26a4[df26a4["业务员"] == "杨阳"]
   df28a4 = df27a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df29a4 = df28a4.rename(columns={'借方': '5月销售'});
   print(df29a4)
   df30a4 = df9[df9["期间"] == 6]  # df63[df63["部门"] == "11部"]
   df31a4 = df30a4[df30a4["业务员"] == "杨阳"]
   df32a4 = df31a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df33a4 = df32a4.rename(columns={'借方': '6月销售'});
   print(df33a4)
   df34a4 = df9[df9["期间"] == 7]  # df63[df63["部门"] == "11部"]
   df35a4 = df34a4[df34a4["业务员"] == "杨阳"]
   df36a4 = df35a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df37a4 = df36a4.rename(columns={'借方': '7月销售'});
   print(df37a4)
   df38a4 = df9[df9["期间"] == 8]  # df63[df63["部门"] == "11部"]
   df39a4 = df38a4[df38a4["业务员"] == "杨阳"]
   df40a4 = df39a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df41a4 = df40a4.rename(columns={'借方': '8月销售'});
   print(df41a4)
   df42a4 = df9[df9["期间"] == 9]  # df63[df63["部门"] == "11部"]
   df43a4 = df42a4[df42a4["业务员"] == "杨阳"]
   df44a4 = df43a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df45a4 = df44a4.rename(columns={'借方': '9月销售'});
   print(df45a4)
   df46a4 = df9[df9["期间"] == 10]  # df63[df63["部门"] == "11部"]
   df47a4 = df46a4[df46a4["业务员"] == "杨阳"]
   df48a4 = df47a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df49a4 = df48a4.rename(columns={'借方': '10月销售'});
   print(df49a4)
   df50a4 = df9[df9["期间"] == 11]  # df63[df63["部门"] == "11部"]
   df51a4 = df50a4[df50a4["业务员"] == "杨阳"]
   df52a4 = df51a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df53a4 = df52a4.rename(columns={'借方': '11月销售'});
   print(df53a4)
   df54a4 = df9[df9["期间"] == 12]  # df63[df63["部门"] == "11部"]
   df55a4 = df54a4[df54a4["业务员"] == "杨阳"]
   df56a4 = df55a4.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df57a4 = df56a4.rename(columns={'借方': '12月销售'});
   print(df57a4)

   df100a4 = pd.merge(df13a4, df17a4, how='left', on=['业务员']);  # 1+2
   df101a4 = pd.merge(df100a4, df21a4, how='left', on=['业务员']);  # +3
   df102a4 = pd.merge(df101a4, df25a4, how='left', on=['业务员']);  # +4
   df103a4 = pd.merge(df102a4, df29a4, how='left', on=['业务员']);  # +5
   df104a4 = pd.merge(df103a4, df33a4, how='left', on=['业务员']);  # +6
   df105a4 = pd.merge(df104a4, df37a4, how='left', on=['业务员']);  # +7
   df106a4 = pd.merge(df105a4, df41a4, how='left', on=['业务员']);  # +8
   df107a4 = pd.merge(df106a4, df45a4, how='left', on=['业务员']);  # +9
   df108a4 = pd.merge(df107a4, df49a4, how='left', on=['业务员']);  # +10
   df109a4 = pd.merge(df108a4, df53a4, how='left', on=['业务员']);  # +11
   df110a4 = pd.merge(df109a4, df57a4, how='left', on=['业务员']);  # +12

   #df1400 = pd.concat([df110aa,df110a1,df110a2,df110a3], ignore_index=True)  # 组合

#####徐蕾
   df10a5 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   df11a5 = df10a5[df10a5["业务员"] == "徐蕾"]
   df12a5 = df11a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df13a5 = df12a5.rename(columns={'借方': '1月销售'});
   print(df13a5)
   df14a5 = df9[df9["期间"] == 2]  # df63[df63["部门"] == "11部"]
   df15a5 = df14a5[df14a5["业务员"] == "徐蕾"]
   df16a5 = df15a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df17a5 = df16a5.rename(columns={'借方': '2月销售'});
   print(df17a5)
   df18a5 = df9[df9["期间"] == 3]  # df63[df63["部门"] == "11部"]
   df19a5 = df18a5[df18a5["业务员"] == "徐蕾"]
   df20a5 = df19a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df21a5 = df20a5.rename(columns={'借方': '3月销售'});
   print(df21a5)
   df22a5 = df9[df9["期间"] == 4]  # df63[df63["部门"] == "11部"]
   df23a5 = df22a5[df22a5["业务员"] == "徐蕾"]
   df24a5 = df23a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df25a5 = df24a5.rename(columns={'借方': '4月销售'});
   print(df25a5)
   df26a5 = df9[df9["期间"] == 5]  # df63[df63["部门"] == "11部"]
   df27a5 = df26a5[df26a5["业务员"] == "徐蕾"]
   df28a5 = df27a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df29a5 = df28a5.rename(columns={'借方': '5月销售'});
   print(df29a5)
   df30a5 = df9[df9["期间"] == 6]  # df63[df63["部门"] == "11部"]
   df31a5 = df30a5[df30a5["业务员"] == "徐蕾"]
   df32a5 = df31a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df33a5 = df32a5.rename(columns={'借方': '6月销售'});
   print(df33a5)
   df34a5 = df9[df9["期间"] == 7]  # df63[df63["部门"] == "11部"]
   df35a5 = df34a5[df34a5["业务员"] == "徐蕾"]
   df36a5 = df35a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df37a5 = df36a5.rename(columns={'借方': '7月销售'});
   print(df37a5)
   df38a5 = df9[df9["期间"] == 8]  # df63[df63["部门"] == "11部"]
   df39a5 = df38a5[df38a5["业务员"] == "徐蕾"]
   df40a5 = df39a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df41a5 = df40a5.rename(columns={'借方': '8月销售'});
   print(df41a5)
   df42a5 = df9[df9["期间"] == 9]  # df63[df63["部门"] == "11部"]
   df43a5 = df42a5[df42a5["业务员"] == "徐蕾"]
   df44a5 = df43a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df45a5 = df44a5.rename(columns={'借方': '9月销售'});
   print(df45a5)
   df46a5 = df9[df9["期间"] == 10]  # df63[df63["部门"] == "11部"]
   df47a5 = df46a5[df46a5["业务员"] == "徐蕾"]
   df48a5 = df47a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df49a5 = df48a5.rename(columns={'借方': '10月销售'});
   print(df49a5)
   df50a5 = df9[df9["期间"] == 11]  # df63[df63["部门"] == "11部"]
   df51a5 = df50a5[df50a5["业务员"] == "徐蕾"]
   df52a5 = df51a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df53a5 = df52a5.rename(columns={'借方': '11月销售'});
   print(df53a5)
   df54a5 = df9[df9["期间"] == 12]  # df63[df63["部门"] == "11部"]
   df55a5 = df54a5[df54a5["业务员"] == "徐蕾"]
   df56a5 = df55a5.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df57a5 = df56a5.rename(columns={'借方': '12月销售'});
   print(df57a5)

   df100a5 = pd.merge(df13a5, df17a5, how='left', on=['业务员']);  # 1+2
   df101a5 = pd.merge(df100a5, df21a5, how='left', on=['业务员']);  # +3
   df102a5 = pd.merge(df101a5, df25a5, how='left', on=['业务员']);  # +4
   df103a5 = pd.merge(df102a5, df29a5, how='left', on=['业务员']);  # +5
   df104a5 = pd.merge(df103a5, df33a5, how='left', on=['业务员']);  # +6
   df105a5 = pd.merge(df104a5, df37a5, how='left', on=['业务员']);  # +7
   df106a5 = pd.merge(df105a5, df41a5, how='left', on=['业务员']);  # +8
   df107a5 = pd.merge(df106a5, df45a5, how='left', on=['业务员']);  # +9
   df108a5 = pd.merge(df107a5, df49a5, how='left', on=['业务员']);  # +10
   df109a5 = pd.merge(df108a5, df53a5, how='left', on=['业务员']);  # +11
   df110a5 = pd.merge(df109a5, df57a5, how='left', on=['业务员']);  # +12

######王伟平

   df10a6 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   df11a6 = df10a6[df10a6["业务员"] == "王伟平"]
   df12a6 = df11a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df13a6 = df12a6.rename(columns={'借方': '1月销售'});
   print(df13a6)
   df14a6 = df9[df9["期间"] == 2]  # df63[df63["部门"] == "11部"]
   df15a6 = df14a6[df14a6["业务员"] == "王伟平"]
   df16a6 = df15a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df17a6 = df16a6.rename(columns={'借方': '2月销售'});
   print(df17a6)
   df18a6 = df9[df9["期间"] == 3]  # df63[df63["部门"] == "11部"]
   df19a6 = df18a6[df18a6["业务员"] == "王伟平"]
   df20a6 = df19a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df21a6 = df20a6.rename(columns={'借方': '3月销售'});
   print(df21a6)
   df22a6 = df9[df9["期间"] == 4]  # df63[df63["部门"] == "11部"]
   df23a6 = df22a6[df22a6["业务员"] == "王伟平"]
   df24a6 = df23a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df25a6 = df24a6.rename(columns={'借方': '4月销售'});
   print(df25a6)
   df26a6 = df9[df9["期间"] == 5]  # df63[df63["部门"] == "11部"]
   df27a6 = df26a6[df26a6["业务员"] == "王伟平"]
   df28a6 = df27a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df29a6 = df28a6.rename(columns={'借方': '5月销售'});
   print(df29a6)
   df30a6 = df9[df9["期间"] == 6]  # df63[df63["部门"] == "11部"]
   df31a6 = df30a6[df30a6["业务员"] == "王伟平"]
   df32a6 = df31a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df33a6 = df32a6.rename(columns={'借方': '6月销售'});
   print(df33a6)
   df34a6 = df9[df9["期间"] == 7]  # df63[df63["部门"] == "11部"]
   df35a6 = df34a6[df34a6["业务员"] == "王伟平"]
   df36a6 = df35a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df37a6 = df36a6.rename(columns={'借方': '7月销售'});
   print(df37a6)
   df38a6 = df9[df9["期间"] == 8]  # df63[df63["部门"] == "11部"]
   df39a6 = df38a6[df38a6["业务员"] == "王伟平"]
   df40a6 = df39a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df41a6 = df40a6.rename(columns={'借方': '8月销售'});
   print(df41a6)
   df42a6 = df9[df9["期间"] == 9]  # df63[df63["部门"] == "11部"]
   df43a6 = df42a6[df42a6["业务员"] == "王伟平"]
   df44a6 = df43a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df45a6 = df44a6.rename(columns={'借方': '9月销售'});
   print(df45a6)
   df46a6 = df9[df9["期间"] == 10]  # df63[df63["部门"] == "11部"]
   df47a6 = df46a6[df46a6["业务员"] == "王伟平"]
   df48a6 = df47a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df49a6 = df48a6.rename(columns={'借方': '10月销售'});
   print(df49a6)
   df50a6 = df9[df9["期间"] == 11]  # df63[df63["部门"] == "11部"]
   df51a6 = df50a6[df50a6["业务员"] == "王伟平"]
   df52a6 = df51a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df53a6 = df52a6.rename(columns={'借方': '11月销售'});
   print(df53a6)
   df54a6 = df9[df9["期间"] == 12]  # df63[df63["部门"] == "11部"]
   df55a6 = df54a6[df54a6["业务员"] == "王伟平"]
   df56a6 = df55a6.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df57a6 = df56a6.rename(columns={'借方': '12月销售'});
   print(df57a6)

   df100a6 = pd.merge(df13a6, df17a6, how='left', on=['业务员']);  # 1+2
   df101a6 = pd.merge(df100a6, df21a6, how='left', on=['业务员']);  # +3
   df102a6 = pd.merge(df101a6, df25a6, how='left', on=['业务员']);  # +4
   df103a6 = pd.merge(df102a6, df29a6, how='left', on=['业务员']);  # +5
   df104a6 = pd.merge(df103a6, df33a6, how='left', on=['业务员']);  # +6
   df105a6 = pd.merge(df104a6, df37a6, how='left', on=['业务员']);  # +7
   df106a6 = pd.merge(df105a6, df41a6, how='left', on=['业务员']);  # +8
   df107a6 = pd.merge(df106a6, df45a6, how='left', on=['业务员']);  # +9
   df108a6 = pd.merge(df107a6, df49a6, how='left', on=['业务员']);  # +10
   df109a6 = pd.merge(df108a6, df53a6, how='left', on=['业务员']);  # +11
   df110a6 = pd.merge(df109a6, df57a6, how='left', on=['业务员']);  # +12

########王宇栋
   df10a7 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   df11a7 = df10a7[df10a7["业务员"] == "王宇栋"]
   df12a7 = df11a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df13a7 = df12a7.rename(columns={'借方': '1月销售'});
   print(df13a7)
   df14a7 = df9[df9["期间"] == 2]  # df63[df63["部门"] == "11部"]
   df15a7 = df14a7[df14a7["业务员"] == "王宇栋"]
   df16a7 = df15a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df17a7 = df16a7.rename(columns={'借方': '2月销售'});
   print(df17a7)
   df18a7 = df9[df9["期间"] == 3]  # df63[df63["部门"] == "11部"]
   df19a7 = df18a7[df18a7["业务员"] == "王宇栋"]
   df20a7 = df19a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df21a7 = df20a7.rename(columns={'借方': '3月销售'});
   print(df21a7)
   df22a7 = df9[df9["期间"] == 4]  # df63[df63["部门"] == "11部"]
   df23a7 = df22a7[df22a7["业务员"] == "王宇栋"]
   df24a7 = df23a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df25a7 = df24a7.rename(columns={'借方': '4月销售'});
   print(df25a7)
   df26a7 = df9[df9["期间"] == 5]  # df63[df63["部门"] == "11部"]
   df27a7 = df26a7[df26a7["业务员"] == "王宇栋"]
   df28a7 = df27a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df29a7 = df28a7.rename(columns={'借方': '5月销售'});
   print(df29a7)
   df30a7 = df9[df9["期间"] == 6]  # df63[df63["部门"] == "11部"]
   df31a7 = df30a7[df30a7["业务员"] == "王宇栋"]
   df32a7 = df31a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df33a7 = df32a7.rename(columns={'借方': '6月销售'});
   print(df33a7)
   df34a7 = df9[df9["期间"] == 7]  # df63[df63["部门"] == "11部"]
   df35a7 = df34a7[df34a7["业务员"] == "王宇栋"]
   df36a7 = df35a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df37a7 = df36a7.rename(columns={'借方': '7月销售'});
   print(df37a7)
   df38a7 = df9[df9["期间"] == 8]  # df63[df63["部门"] == "11部"]
   df39a7 = df38a7[df38a7["业务员"] == "王宇栋"]
   df40a7 = df39a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df41a7 = df40a7.rename(columns={'借方': '8月销售'});
   print(df41a7)
   df42a7 = df9[df9["期间"] == 9]  # df63[df63["部门"] == "11部"]
   df43a7 = df42a7[df42a7["业务员"] == "王宇栋"]
   df44a7 = df43a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df45a7 = df44a7.rename(columns={'借方': '9月销售'});
   print(df45a7)
   df46a7 = df9[df9["期间"] == 10]  # df63[df63["部门"] == "11部"]
   df47a7 = df46a7[df46a7["业务员"] == "王宇栋"]
   df48a7 = df47a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df49a7 = df48a7.rename(columns={'借方': '10月销售'});
   print(df49a7)
   df50a7 = df9[df9["期间"] == 11]  # df63[df63["部门"] == "11部"]
   df51a7 = df50a7[df50a7["业务员"] == "王宇栋"]
   df52a7 = df51a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df53a7 = df52a7.rename(columns={'借方': '11月销售'});
   print(df53a7)
   df54a7 = df9[df9["期间"] == 12]  # df63[df63["部门"] == "11部"]
   df55a7 = df54a7[df54a7["业务员"] == "王宇栋"]
   df56a7 = df55a7.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df57a7 = df56a7.rename(columns={'借方': '12月销售'});
   print(df57a7)

   df100a7 = pd.merge(df13a7, df17a7, how='left', on=['业务员']);  # 1+2
   df101a7 = pd.merge(df100a7, df21a7, how='left', on=['业务员']);  # +3
   df102a7 = pd.merge(df101a7, df25a7, how='left', on=['业务员']);  # +4
   df103a7 = pd.merge(df102a7, df29a7, how='left', on=['业务员']);  # +5
   df104a7 = pd.merge(df103a7, df33a7, how='left', on=['业务员']);  # +6
   df105a7 = pd.merge(df104a7, df37a7, how='left', on=['业务员']);  # +7
   df106a7 = pd.merge(df105a7, df41a7, how='left', on=['业务员']);  # +8
   df107a7 = pd.merge(df106a7, df45a7, how='left', on=['业务员']);  # +9
   df108a7 = pd.merge(df107a7, df49a7, how='left', on=['业务员']);  # +10
   df109a7 = pd.merge(df108a7, df53a7, how='left', on=['业务员']);  # +11
   df110a7 = pd.merge(df109a7, df57a7, how='left', on=['业务员']);  # +12
#####朱津齐
   df10a8 = df9[df9["期间"] == 1]  # df63[df63["部门"] == "11部"]
   df11a8 = df10a8[df10a8["业务员"] == "朱津齐"]
   df12a8 = df11a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df13a8 = df12a8.rename(columns={'借方': '1月销售'});
   print(df13a8)
   df14a8 = df9[df9["期间"] == 2]  # df63[df63["部门"] == "11部"]
   df15a8 = df14a8[df14a8["业务员"] == "朱津齐"]
   df16a8 = df15a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df17a8 = df16a8.rename(columns={'借方': '2月销售'});
   print(df17a8)
   df18a8 = df9[df9["期间"] == 3]  # df63[df63["部门"] == "11部"]
   df19a8 = df18a8[df18a8["业务员"] == "朱津齐"]
   df20a8 = df19a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df21a8 = df20a8.rename(columns={'借方': '3月销售'});
   print(df21a8)
   df22a8 = df9[df9["期间"] == 4]  # df63[df63["部门"] == "11部"]
   df23a8 = df22a8[df22a8["业务员"] == "朱津齐"]
   df24a8 = df23a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df25a8 = df24a8.rename(columns={'借方': '4月销售'});
   print(df25a8)
   df26a8 = df9[df9["期间"] == 5]  # df63[df63["部门"] == "11部"]
   df27a8 = df26a8[df26a8["业务员"] == "朱津齐"]
   df28a8 = df27a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df29a8 = df28a8.rename(columns={'借方': '5月销售'});
   print(df29a8)
   df30a8 = df9[df9["期间"] == 6]  # df63[df63["部门"] == "11部"]
   df31a8 = df30a8[df30a8["业务员"] == "朱津齐"]
   df32a8 = df31a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df33a8 = df32a8.rename(columns={'借方': '6月销售'});
   print(df33a8)
   df34a8 = df9[df9["期间"] == 7]  # df63[df63["部门"] == "11部"]
   df35a8 = df34a8[df34a8["业务员"] == "朱津齐"]
   df36a8 = df35a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df37a8 = df36a8.rename(columns={'借方': '7月销售'});
   print(df37a8)
   df38a8 = df9[df9["期间"] == 8]  # df63[df63["部门"] == "11部"]
   df39a8 = df38a8[df38a8["业务员"] == "朱津齐"]
   df40a8 = df39a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df41a8 = df40a8.rename(columns={'借方': '8月销售'});
   print(df41a8)
   df42a8 = df9[df9["期间"] == 9]  # df63[df63["部门"] == "11部"]
   df43a8 = df42a8[df42a8["业务员"] == "朱津齐"]
   df44a8 = df43a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df45a8 = df44a8.rename(columns={'借方': '9月销售'});
   print(df45a8)
   df46a8 = df9[df9["期间"] == 10]  # df63[df63["部门"] == "11部"]
   df47a8 = df46a8[df46a8["业务员"] == "朱津齐"]
   df48a8 = df47a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df49a8 = df48a8.rename(columns={'借方': '10月销售'});
   print(df49a8)
   df50a8 = df9[df9["期间"] == 11]  # df63[df63["部门"] == "11部"]
   df51a8 = df50a8[df50a8["业务员"] == "朱津齐"]
   df52a8 = df51a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df53a8 = df52a8.rename(columns={'借方': '11月销售'});
   print(df53a8)
   df54a8 = df9[df9["期间"] == 12]  # df63[df63["部门"] == "11部"]
   df55a8 = df54a8[df54a8["业务员"] == "朱津齐"]
   df56a8 = df55a8.drop(["期间", "期初余额", "本年累计借方", "贷方", "本年累计贷方", "期末余额", "年初余额"], axis=1)  # 删列
   df57a8 = df56a8.rename(columns={'借方': '12月销售'});
   print(df57a8)

   df100a8 = pd.merge(df13a8, df17a8, how='left', on=['业务员']);  # 1+2
   df101a8 = pd.merge(df100a8, df21a8, how='left', on=['业务员']);  # +3
   df102a8 = pd.merge(df101a8, df25a8, how='left', on=['业务员']);  # +4
   df103a8 = pd.merge(df102a8, df29a8, how='left', on=['业务员']);  # +5
   df104a8 = pd.merge(df103a8, df33a8, how='left', on=['业务员']);  # +6
   df105a8 = pd.merge(df104a8, df37a8, how='left', on=['业务员']);  # +7
   df106a8 = pd.merge(df105a8, df41a8, how='left', on=['业务员']);  # +8
   df107a8 = pd.merge(df106a8, df45a8, how='left', on=['业务员']);  # +9
   df108a8 = pd.merge(df107a8, df49a8, how='left', on=['业务员']);  # +10
   df109a8 = pd.merge(df108a8, df53a8, how='left', on=['业务员']);  # +11
   df110a8 = pd.merge(df109a8, df57a8, how='left', on=['业务员']);  # +12

   #####药品二部销售小计
   df1100aa = pd.concat([df110a4, df110a5, df110a6,df110a7,df110a8], ignore_index=True)  # 组合

   df1100aa["部门"] = df1100aa["业务员"]
   # df1001=df1000.rename(columns={'傅诗云': '药品1部','戚肆朝': '药品1部','张丽君': '药品1部','王振杰': '药品1部'});

   df1100aa["部门"].replace("杨阳", "药品2部", inplace=True)
   df1100aa["部门"].replace("徐蕾", "药品2部", inplace=True)
   df1100aa["部门"].replace("王伟平", "药品2部", inplace=True)
   df1100aa["部门"].replace("王宇栋", "药品2部", inplace=True)
   df1100aa["部门"].replace("朱津齐", "药品2部", inplace=True)

   df1101aa = df1100aa.groupby(["部门"], as_index=False)[
       "1月销售", "2月销售", "3月销售", "4月销售", "5月销售", "6月销售", "7月销售", "8月销售", "9月销售", "10月销售",
       "11月销售", "12月销售"].sum();

   df1102aa = pd.concat([df1100aa, df1101aa], ignore_index=True)  # 组合

   df1103aa = df1102aa.fillna(0)
   df1103aa["业务员"].replace(0, "药品2部", inplace=True)

   print(df1103aa)

   df1104aa = df1103aa.groupby(["业务员"], as_index=False)[
       "1月销售", "2月销售", "3月销售", "4月销售", "5月销售", "6月销售", "7月销售", "8月销售", "9月销售", "10月销售",
       "11月销售", "12月销售"].sum();

   df1400 = pd.concat([df1104a, df1104aa], ignore_index=True)  # 组合



   #####表4平均销售

   df201ab = df1400[df1400["业务员"] == "傅诗云"]
   df201ab["1月平均销售"] = df201ab["1月销售"]
   df201ab["2月平均销售"] = (df201ab["1月销售"] + df201ab["2月销售"]) / 2
   df201ab["3月平均销售"] = (df201ab["1月销售"] + df201ab["2月销售"] + df201ab["3月销售"]) / 3
   df201ab["4月平均销售"] = (df201ab["1月销售"] + df201ab["2月销售"] + df201ab["3月销售"] + df201ab["4月销售"]) / 4
   df201ab["5月平均销售"] = (df201ab["1月销售"] + df201ab["2月销售"] + df201ab["3月销售"] + df201ab[
       "4月销售"] + df201ab["5月销售"]) / 5
   df201ab["6月平均销售"] = (df201ab["1月销售"] + df201ab["2月销售"] + df201ab["3月销售"] + df201ab[
       "4月销售"] +
                        df201ab["5月销售"] + df201ab["6月销售"]) / 6
   df201ab["7月平均销售"] = (df201ab["1月销售"] + df201ab["2月销售"] + df201ab["3月销售"] + df201ab[
       "4月销售"] +
                        df201ab["5月销售"] + df201ab["6月销售"] + df201ab["7月销售"]) / 7
   df201ab["8月平均销售"] = (df201ab["1月销售"] + df201ab["2月销售"] + df201ab["3月销售"] + df201ab[
       "4月销售"] +
                        df201ab["5月销售"] + df201ab["6月销售"] + df201ab["7月销售"] + df201ab["8月销售"]) / 8
   df201ab["9月平均销售"] = (df201ab["1月销售"] + df201ab["2月销售"] + df201ab["3月销售"] + df201ab[
       "4月销售"] +
                        df201ab["5月销售"] + df201ab["6月销售"] + df201ab["7月销售"] + df201ab["8月销售"] + df201ab[
                            "9月销售"]) / 9
   df201ab["10月平均销售"] = (df201ab["1月销售"]  + df201ab["2月销售"] + df201ab["3月销售"] + df201ab[
       "4月销售"] +
                         df201ab["5月销售"] + df201ab["6月销售"] + df201ab["7月销售"] + df201ab["8月销售"] + df201ab[
                             "9月销售"] + df201ab["10月销售"]) / 10
   df201ab["11月平均销售"] = (df201ab["1月销售"]  + df201ab["2月销售"] + df201ab["3月销售"] + df201ab[
       "4月销售"] +
                         df201ab["5月销售"] + df201ab["6月销售"] + df201ab["7月销售"] + df201ab["8月销售"] + df201ab[
                             "9月销售"] + df201ab["10月销售"] + df201ab["11月销售"]) / 11
   df201ab["12月平均销售"] = (df201ab["1月销售"]  + df201ab["2月销售"] + df201ab["3月销售"] + df201ab[
       "4月销售"] +
                         df201ab["5月销售"] + df201ab["6月销售"] + df201ab["7月销售"] + df201ab["8月销售"] + df201ab[
                             "9月销售"] + df201ab["10月销售"] + df201ab["11月销售"] + df201ab["12月销售"]) / 12

   df201ab1 = df1400[df1400["业务员"] == "戚肆朝"]
   df201ab1["1月平均销售"] = df201ab1["1月销售"]
   df201ab1["2月平均销售"] = (df201ab1["1月销售"] + df201ab1["2月销售"]) / 2
   df201ab1["3月平均销售"] = (df201ab1["1月销售"] + df201ab1["2月销售"] + df201ab1["3月销售"]) / 3
   df201ab1["4月平均销售"] = (df201ab1["1月销售"] + df201ab1["2月销售"] + df201ab1["3月销售"] + df201ab1["4月销售"]) / 4
   df201ab1["5月平均销售"] = (df201ab1["1月销售"] + df201ab1["2月销售"] + df201ab1["3月销售"] + df201ab1[
       "4月销售"] + df201ab1["5月销售"]) / 5
   df201ab1["6月平均销售"] = (df201ab1["1月销售"] + df201ab1["2月销售"] + df201ab1["3月销售"] + df201ab1[
       "4月销售"] +
                         df201ab1["5月销售"] + df201ab1["6月销售"]) / 6
   df201ab1["7月平均销售"] = (df201ab1["1月销售"] + df201ab1["2月销售"] + df201ab1["3月销售"] + df201ab1[
       "4月销售"] +
                         df201ab1["5月销售"] + df201ab1["6月销售"] + df201ab1["7月销售"]) / 7
   df201ab1["8月平均销售"] = (df201ab1["1月销售"] + df201ab1["2月销售"] + df201ab1["3月销售"] + df201ab1[
       "4月销售"] +
                         df201ab1["5月销售"] + df201ab1["6月销售"] + df201ab1["7月销售"] + df201ab1["8月销售"]) / 8
   df201ab1["9月平均销售"] = (df201ab1["1月销售"] + df201ab1["2月销售"] + df201ab1["3月销售"] + df201ab1[
       "4月销售"] +
                         df201ab1["5月销售"] + df201ab1["6月销售"] + df201ab1["7月销售"] + df201ab1["8月销售"] + df201ab1[
                             "9月销售"]) / 9
   df201ab1["10月平均销售"] = (df201ab1["1月销售"] + df201ab1["2月销售"] + df201ab1["3月销售"] + df201ab1[
       "4月销售"] +
                          df201ab1["5月销售"] + df201ab1["6月销售"] + df201ab1["7月销售"] + df201ab1["8月销售"] + df201ab1[
                              "9月销售"] + df201ab1["10月销售"]) / 10
   df201ab1["11月平均销售"] = (df201ab1["1月销售"] + df201ab1["2月销售"] + df201ab1["3月销售"] + df201ab1[
       "4月销售"] +
                          df201ab1["5月销售"] + df201ab1["6月销售"] + df201ab1["7月销售"] + df201ab1["8月销售"] + df201ab1[
                              "9月销售"] + df201ab1["10月销售"] + df201ab1["11月销售"]) / 11
   df201ab1["12月平均销售"] = (df201ab1["1月销售"] + df201ab1["2月销售"] + df201ab1["3月销售"] + df201ab1[
       "4月销售"] +
                          df201ab1["5月销售"] + df201ab1["6月销售"] + df201ab1["7月销售"] + df201ab1["8月销售"] + df201ab1[
                              "9月销售"] + df201ab1["10月销售"] + df201ab1["11月销售"] + df201ab1["12月销售"]) / 12
   #print(df201ab)
   df201ab2 = df1400[df1400["业务员"] == "张丽君"]
   df201ab2["1月平均销售"] = df201ab2["1月销售"]
   df201ab2["2月平均销售"] = (df201ab2["1月销售"] + df201ab2["2月销售"]) / 2
   df201ab2["3月平均销售"] = (df201ab2["1月销售"] + df201ab2["2月销售"] + df201ab2["3月销售"]) / 3
   df201ab2["4月平均销售"] = (df201ab2["1月销售"] + df201ab2["2月销售"] + df201ab2["3月销售"] + df201ab2["4月销售"]) / 4
   df201ab2["5月平均销售"] = (df201ab2["1月销售"] + df201ab2["2月销售"] + df201ab2["3月销售"] + df201ab2[
       "4月销售"] + df201ab2["5月销售"]) / 5
   df201ab2["6月平均销售"] = (df201ab2["1月销售"] + df201ab2["2月销售"] + df201ab2["3月销售"] + df201ab2[
       "4月销售"] +
                         df201ab2["5月销售"] + df201ab2["6月销售"]) / 6
   df201ab2["7月平均销售"] = (df201ab2["1月销售"] + df201ab2["2月销售"] + df201ab2["3月销售"] + df201ab2[
       "4月销售"] +
                         df201ab2["5月销售"] + df201ab2["6月销售"] + df201ab2["7月销售"]) / 7
   df201ab2["8月平均销售"] = (df201ab2["1月销售"] + df201ab2["2月销售"] + df201ab2["3月销售"] + df201ab2[
       "4月销售"] +
                         df201ab2["5月销售"] + df201ab2["6月销售"] + df201ab2["7月销售"] + df201ab2["8月销售"]) / 8
   df201ab2["9月平均销售"] = (df201ab2["1月销售"] + df201ab2["2月销售"] + df201ab2["3月销售"] + df201ab2[
       "4月销售"] +
                         df201ab2["5月销售"] + df201ab2["6月销售"] + df201ab2["7月销售"] + df201ab2["8月销售"] + df201ab2[
                             "9月销售"]) / 9
   df201ab2["10月平均销售"] = (df201ab2["1月销售"] + df201ab2["2月销售"] + df201ab2["3月销售"] + df201ab2[
       "4月销售"] +
                          df201ab2["5月销售"] + df201ab2["6月销售"] + df201ab2["7月销售"] + df201ab2["8月销售"] + df201ab2[
                              "9月销售"] + df201ab2["10月销售"]) / 10
   df201ab2["11月平均销售"] = (df201ab2["1月销售"] + df201ab2["2月销售"] + df201ab2["3月销售"] + df201ab2[
       "4月销售"] +
                          df201ab2["5月销售"] + df201ab2["6月销售"] + df201ab2["7月销售"] + df201ab2["8月销售"] + df201ab2[
                              "9月销售"] + df201ab2["10月销售"] + df201ab2["11月销售"]) / 11
   df201ab2["12月平均销售"] = (df201ab2["1月销售"] + df201ab2["2月销售"] + df201ab2["3月销售"] + df201ab2[
       "4月销售"] +
                          df201ab2["5月销售"] + df201ab2["6月销售"] + df201ab2["7月销售"] + df201ab2["8月销售"] + df201ab2[
                              "9月销售"] + df201ab2["10月销售"] + df201ab2["11月销售"] + df201ab2["12月销售"]) / 12

   df201ab3 = df1400[df1400["业务员"] == "王振杰"]
   df201ab3["1月平均销售"] = df201ab3["1月销售"]
   df201ab3["2月平均销售"] = (df201ab3["1月销售"] + df201ab3["2月销售"]) / 2
   df201ab3["3月平均销售"] = (df201ab3["1月销售"] + df201ab3["2月销售"] + df201ab3["3月销售"]) / 3
   df201ab3["4月平均销售"] = (df201ab3["1月销售"] + df201ab3["2月销售"] + df201ab3["3月销售"] + df201ab3["4月销售"]) / 4
   df201ab3["5月平均销售"] = (df201ab3["1月销售"] + df201ab3["2月销售"] + df201ab3["3月销售"] + df201ab3[
       "4月销售"] + df201ab3["5月销售"]) / 5
   df201ab3["6月平均销售"] = (df201ab3["1月销售"] + df201ab3["2月销售"] + df201ab3["3月销售"] + df201ab3[
       "4月销售"] +
                         df201ab3["5月销售"] + df201ab3["6月销售"]) / 6
   df201ab3["7月平均销售"] = (df201ab3["1月销售"] + df201ab3["2月销售"] + df201ab3["3月销售"] + df201ab3[
       "4月销售"] +
                         df201ab3["5月销售"] + df201ab3["6月销售"] + df201ab3["7月销售"]) / 7
   df201ab3["8月平均销售"] = (df201ab3["1月销售"] + df201ab3["2月销售"] + df201ab3["3月销售"] + df201ab3[
       "4月销售"] +
                         df201ab3["5月销售"] + df201ab3["6月销售"] + df201ab3["7月销售"] + df201ab3["8月销售"]) / 8
   df201ab3["9月平均销售"] = (df201ab3["1月销售"] + df201ab3["2月销售"] + df201ab3["3月销售"] + df201ab3[
       "4月销售"] +
                         df201ab3["5月销售"] + df201ab3["6月销售"] + df201ab3["7月销售"] + df201ab3["8月销售"] + df201ab3[
                             "9月销售"]) / 9
   df201ab3["10月平均销售"] = (df201ab3["1月销售"] + df201ab3["2月销售"] + df201ab3["3月销售"] + df201ab3[
       "4月销售"] +
                          df201ab3["5月销售"] + df201ab3["6月销售"] + df201ab3["7月销售"] + df201ab3["8月销售"] + df201ab3[
                              "9月销售"] + df201ab3["10月销售"]) / 10
   df201ab3["11月平均销售"] = (df201ab3["1月销售"] + df201ab3["2月销售"] + df201ab3["3月销售"] + df201ab3[
       "4月销售"] +
                          df201ab3["5月销售"] + df201ab3["6月销售"] + df201ab3["7月销售"] + df201ab3["8月销售"] + df201ab3[
                              "9月销售"] + df201ab3["10月销售"] + df201ab3["11月销售"]) / 11
   df201ab3["12月平均销售"] = (df201ab3["1月销售"] + df201ab3["2月销售"] + df201ab3["3月销售"] + df201ab3[
       "4月销售"] +
                          df201ab3["5月销售"] + df201ab3["6月销售"] + df201ab3["7月销售"] + df201ab3["8月销售"] + df201ab3[
                              "9月销售"] + df201ab3["10月销售"] + df201ab3["11月销售"] + df201ab3["12月销售"]) / 12

   df201ab4 = df1400[df1400["业务员"] == "药品1部"]
   df201ab4["1月平均销售"] = df201ab4["1月销售"]
   df201ab4["2月平均销售"] = (df201ab4["1月销售"] + df201ab4["2月销售"]) / 2
   df201ab4["3月平均销售"] = (df201ab4["1月销售"] + df201ab4["2月销售"] + df201ab4["3月销售"]) / 3
   df201ab4["4月平均销售"] = (df201ab4["1月销售"] + df201ab4["2月销售"] + df201ab4["3月销售"] + df201ab4["4月销售"]) / 4
   df201ab4["5月平均销售"] = (df201ab4["1月销售"] + df201ab4["2月销售"] + df201ab4["3月销售"] + df201ab4[
       "4月销售"] + df201ab4["5月销售"]) / 5
   df201ab4["6月平均销售"] = (df201ab4["1月销售"] + df201ab4["2月销售"] + df201ab4["3月销售"] + df201ab4[
       "4月销售"] +
                         df201ab4["5月销售"] + df201ab4["6月销售"]) / 6
   df201ab4["7月平均销售"] = (df201ab4["1月销售"] + df201ab4["2月销售"] + df201ab4["3月销售"] + df201ab4[
       "4月销售"] +
                         df201ab4["5月销售"] + df201ab4["6月销售"] + df201ab4["7月销售"]) / 7
   df201ab4["8月平均销售"] = (df201ab4["1月销售"] + df201ab4["2月销售"] + df201ab4["3月销售"] + df201ab4[
       "4月销售"] +
                         df201ab4["5月销售"] + df201ab4["6月销售"] + df201ab4["7月销售"] + df201ab4["8月销售"]) / 8
   df201ab4["9月平均销售"] = (df201ab4["1月销售"] + df201ab4["2月销售"] + df201ab4["3月销售"] + df201ab4[
       "4月销售"] +
                         df201ab4["5月销售"] + df201ab4["6月销售"] + df201ab4["7月销售"] + df201ab4["8月销售"] + df201ab4[
                             "9月销售"]) / 9
   df201ab4["10月平均销售"] = (df201ab4["1月销售"] + df201ab4["2月销售"] + df201ab4["3月销售"] + df201ab4[
       "4月销售"] +
                          df201ab4["5月销售"] + df201ab4["6月销售"] + df201ab4["7月销售"] + df201ab4["8月销售"] + df201ab4[
                              "9月销售"] + df201ab4["10月销售"]) / 10
   df201ab4["11月平均销售"] = (df201ab4["1月销售"] + df201ab4["2月销售"] + df201ab4["3月销售"] + df201ab4[
       "4月销售"] +
                          df201ab4["5月销售"] + df201ab4["6月销售"] + df201ab4["7月销售"] + df201ab4["8月销售"] + df201ab4[
                              "9月销售"] + df201ab4["10月销售"] + df201ab4["11月销售"]) / 11
   df201ab4["12月平均销售"] = (df201ab4["1月销售"] + df201ab4["2月销售"] + df201ab4["3月销售"] + df201ab4[
       "4月销售"] +
                          df201ab4["5月销售"] + df201ab4["6月销售"] + df201ab4["7月销售"] + df201ab4["8月销售"] + df201ab4[
                              "9月销售"] + df201ab4["10月销售"] + df201ab4["11月销售"] + df201ab4["12月销售"]) / 12

   df201ab5 = df1400[df1400["业务员"] == "杨阳"]
   df201ab5["1月平均销售"] = df201ab5["1月销售"]
   df201ab5["2月平均销售"] = (df201ab5["1月销售"] + df201ab5["2月销售"]) / 2
   df201ab5["3月平均销售"] = (df201ab5["1月销售"] + df201ab5["2月销售"] + df201ab5["3月销售"]) / 3
   df201ab5["4月平均销售"] = (df201ab5["1月销售"] + df201ab5["2月销售"] + df201ab5["3月销售"] + df201ab5["4月销售"]) / 4
   df201ab5["5月平均销售"] = (df201ab5["1月销售"] + df201ab5["2月销售"] + df201ab5["3月销售"] + df201ab5[
       "4月销售"] + df201ab5["5月销售"]) / 5
   df201ab5["6月平均销售"] = (df201ab5["1月销售"] + df201ab5["2月销售"] + df201ab5["3月销售"] + df201ab5[
       "4月销售"] +
                         df201ab5["5月销售"] + df201ab5["6月销售"]) / 6
   df201ab5["7月平均销售"] = (df201ab5["1月销售"] + df201ab5["2月销售"] + df201ab5["3月销售"] + df201ab5[
       "4月销售"] +
                         df201ab5["5月销售"] + df201ab5["6月销售"] + df201ab5["7月销售"]) / 7
   df201ab5["8月平均销售"] = (df201ab5["1月销售"] + df201ab5["2月销售"] + df201ab5["3月销售"] + df201ab5[
       "4月销售"] +
                         df201ab5["5月销售"] + df201ab5["6月销售"] + df201ab5["7月销售"] + df201ab5["8月销售"]) / 8
   df201ab5["9月平均销售"] = (df201ab5["1月销售"] + df201ab5["2月销售"] + df201ab5["3月销售"] + df201ab5[
       "4月销售"] +
                         df201ab5["5月销售"] + df201ab5["6月销售"] + df201ab5["7月销售"] + df201ab5["8月销售"] + df201ab5[
                             "9月销售"]) / 9
   df201ab5["10月平均销售"] = (df201ab5["1月销售"] + df201ab5["2月销售"] + df201ab5["3月销售"] + df201ab5[
       "4月销售"] +
                          df201ab5["5月销售"] + df201ab5["6月销售"] + df201ab5["7月销售"] + df201ab5["8月销售"] + df201ab5[
                              "9月销售"] + df201ab5["10月销售"]) / 10
   df201ab5["11月平均销售"] = (df201ab5["1月销售"] + df201ab5["2月销售"] + df201ab5["3月销售"] + df201ab5[
       "4月销售"] +
                          df201ab5["5月销售"] + df201ab5["6月销售"] + df201ab5["7月销售"] + df201ab5["8月销售"] + df201ab5[
                              "9月销售"] + df201ab5["10月销售"] + df201ab5["11月销售"]) / 11
   df201ab5["12月平均销售"] = (df201ab5["1月销售"] + df201ab5["2月销售"] + df201ab5["3月销售"] + df201ab5[
       "4月销售"] +
                          df201ab5["5月销售"] + df201ab5["6月销售"] + df201ab5["7月销售"] + df201ab5["8月销售"] + df201ab5[
                              "9月销售"] + df201ab5["10月销售"] + df201ab5["11月销售"] + df201ab5["12月销售"]) / 12
   df201ab6 = df1400[df1400["业务员"] == "徐蕾"]
   df201ab6["1月平均销售"] = df201ab6["1月销售"]
   df201ab6["2月平均销售"] = (df201ab6["1月销售"] + df201ab6["2月销售"]) / 2
   df201ab6["3月平均销售"] = (df201ab6["1月销售"] + df201ab6["2月销售"] + df201ab6["3月销售"]) / 3
   df201ab6["4月平均销售"] = (df201ab6["1月销售"] + df201ab6["2月销售"] + df201ab6["3月销售"] + df201ab6["4月销售"]) / 4
   df201ab6["5月平均销售"] = (df201ab6["1月销售"] + df201ab6["2月销售"] + df201ab6["3月销售"] + df201ab6[
       "4月销售"] + df201ab6["5月销售"]) / 5
   df201ab6["6月平均销售"] = (df201ab6["1月销售"] + df201ab6["2月销售"] + df201ab6["3月销售"] + df201ab6[
       "4月销售"] +
                         df201ab6["5月销售"] + df201ab6["6月销售"]) / 6
   df201ab6["7月平均销售"] = (df201ab6["1月销售"] + df201ab6["2月销售"] + df201ab6["3月销售"] + df201ab6[
       "4月销售"] +
                         df201ab6["5月销售"] + df201ab6["6月销售"] + df201ab6["7月销售"]) / 7
   df201ab6["8月平均销售"] = (df201ab6["1月销售"] + df201ab6["2月销售"] + df201ab6["3月销售"] + df201ab6[
       "4月销售"] +
                         df201ab6["5月销售"] + df201ab6["6月销售"] + df201ab6["7月销售"] + df201ab6["8月销售"]) / 8
   df201ab6["9月平均销售"] = (df201ab6["1月销售"] + df201ab6["2月销售"] + df201ab6["3月销售"] + df201ab6[
       "4月销售"] +
                         df201ab6["5月销售"] + df201ab6["6月销售"] + df201ab6["7月销售"] + df201ab6["8月销售"] + df201ab6[
                             "9月销售"]) / 9
   df201ab6["10月平均销售"] = (df201ab6["1月销售"] + df201ab6["2月销售"] + df201ab6["3月销售"] + df201ab6[
       "4月销售"] +
                          df201ab6["5月销售"] + df201ab6["6月销售"] + df201ab6["7月销售"] + df201ab6["8月销售"] + df201ab6[
                              "9月销售"] + df201ab6["10月销售"]) / 10
   df201ab6["11月平均销售"] = (df201ab6["1月销售"] + df201ab6["2月销售"] + df201ab6["3月销售"] + df201ab6[
       "4月销售"] +
                          df201ab6["5月销售"] + df201ab6["6月销售"] + df201ab6["7月销售"] + df201ab6["8月销售"] + df201ab6[
                              "9月销售"] + df201ab6["10月销售"] + df201ab6["11月销售"]) / 11
   df201ab6["12月平均销售"] = (df201ab6["1月销售"] + df201ab6["2月销售"] + df201ab6["3月销售"] + df201ab6[
       "4月销售"] +
                          df201ab6["5月销售"] + df201ab6["6月销售"] + df201ab6["7月销售"] + df201ab6["8月销售"] + df201ab6[
                              "9月销售"] + df201ab6["10月销售"] + df201ab6["11月销售"] + df201ab6["12月销售"]) / 12
   df201ab7 = df1400[df1400["业务员"] == "王伟平"]
   df201ab7["1月平均销售"] = df201ab7["1月销售"]
   df201ab7["2月平均销售"] = (df201ab7["1月销售"] + df201ab7["2月销售"]) / 2
   df201ab7["3月平均销售"] = (df201ab7["1月销售"] + df201ab7["2月销售"] + df201ab7["3月销售"]) / 3
   df201ab7["4月平均销售"] = (df201ab7["1月销售"] + df201ab7["2月销售"] + df201ab7["3月销售"] + df201ab7["4月销售"]) / 4
   df201ab7["5月平均销售"] = (df201ab7["1月销售"] + df201ab7["2月销售"] + df201ab7["3月销售"] + df201ab7[
       "4月销售"] + df201ab7["5月销售"]) / 5
   df201ab7["6月平均销售"] = (df201ab7["1月销售"] + df201ab7["2月销售"] + df201ab7["3月销售"] + df201ab7[
       "4月销售"] +
                         df201ab7["5月销售"] + df201ab7["6月销售"]) / 6
   df201ab7["7月平均销售"] = (df201ab7["1月销售"] + df201ab7["2月销售"] + df201ab7["3月销售"] + df201ab7[
       "4月销售"] +
                         df201ab7["5月销售"] + df201ab7["6月销售"] + df201ab7["7月销售"]) / 7
   df201ab7["8月平均销售"] = (df201ab7["1月销售"] + df201ab7["2月销售"] + df201ab7["3月销售"] + df201ab7[
       "4月销售"] +
                         df201ab7["5月销售"] + df201ab7["6月销售"] + df201ab7["7月销售"] + df201ab7["8月销售"]) / 8
   df201ab7["9月平均销售"] = (df201ab7["1月销售"] + df201ab7["2月销售"] + df201ab7["3月销售"] + df201ab7[
       "4月销售"] +
                         df201ab7["5月销售"] + df201ab7["6月销售"] + df201ab7["7月销售"] + df201ab7["8月销售"] + df201ab7[
                             "9月销售"]) / 9
   df201ab7["10月平均销售"] = (df201ab7["1月销售"] + df201ab7["2月销售"] + df201ab7["3月销售"] + df201ab7[
       "4月销售"] +
                          df201ab7["5月销售"] + df201ab7["6月销售"] + df201ab7["7月销售"] + df201ab7["8月销售"] + df201ab7[
                              "9月销售"] + df201ab7["10月销售"]) / 10
   df201ab7["11月平均销售"] = (df201ab7["1月销售"] + df201ab7["2月销售"] + df201ab7["3月销售"] + df201ab7[
       "4月销售"] +
                          df201ab7["5月销售"] + df201ab7["6月销售"] + df201ab7["7月销售"] + df201ab7["8月销售"] + df201ab7[
                              "9月销售"] + df201ab7["10月销售"] + df201ab7["11月销售"]) / 11
   df201ab7["12月平均销售"] = (df201ab7["1月销售"] + df201ab7["2月销售"] + df201ab7["3月销售"] + df201ab7[
       "4月销售"] +
                          df201ab7["5月销售"] + df201ab7["6月销售"] + df201ab7["7月销售"] + df201ab7["8月销售"] + df201ab7[
                              "9月销售"] + df201ab7["10月销售"] + df201ab7["11月销售"] + df201ab7["12月销售"]) / 12

   df201ab8 = df1400[df1400["业务员"] == "王宇栋"]
   df201ab8["1月平均销售"] = df201ab8["1月销售"]
   df201ab8["2月平均销售"] = (df201ab8["1月销售"] + df201ab8["2月销售"]) / 2
   df201ab8["3月平均销售"] = (df201ab8["1月销售"] + df201ab8["2月销售"] + df201ab8["3月销售"]) / 3
   df201ab8["4月平均销售"] = (df201ab8["1月销售"] + df201ab8["2月销售"] + df201ab8["3月销售"] + df201ab8["4月销售"]) / 4
   df201ab8["5月平均销售"] = (df201ab8["1月销售"] + df201ab8["2月销售"] + df201ab8["3月销售"] + df201ab8[
       "4月销售"] + df201ab8["5月销售"]) / 5
   df201ab8["6月平均销售"] = (df201ab8["1月销售"] + df201ab8["2月销售"] + df201ab8["3月销售"] + df201ab8[
       "4月销售"] +
                         df201ab8["5月销售"] + df201ab8["6月销售"]) / 6
   df201ab8["7月平均销售"] = (df201ab8["1月销售"] + df201ab8["2月销售"] + df201ab8["3月销售"] + df201ab8[
       "4月销售"] +
                         df201ab8["5月销售"] + df201ab8["6月销售"] + df201ab8["7月销售"]) / 7
   df201ab8["8月平均销售"] = (df201ab8["1月销售"] + df201ab8["2月销售"] + df201ab8["3月销售"] + df201ab8[
       "4月销售"] +
                         df201ab8["5月销售"] + df201ab8["6月销售"] + df201ab8["7月销售"] + df201ab8["8月销售"]) / 8
   df201ab8["9月平均销售"] = (df201ab8["1月销售"] + df201ab8["2月销售"] + df201ab8["3月销售"] + df201ab8[
       "4月销售"] +
                         df201ab8["5月销售"] + df201ab8["6月销售"] + df201ab8["7月销售"] + df201ab8["8月销售"] + df201ab8[
                             "9月销售"]) / 9
   df201ab8["10月平均销售"] = (df201ab8["1月销售"] + df201ab8["2月销售"] + df201ab8["3月销售"] + df201ab8[
       "4月销售"] +
                          df201ab8["5月销售"] + df201ab8["6月销售"] + df201ab8["7月销售"] + df201ab8["8月销售"] + df201ab8[
                              "9月销售"] + df201ab8["10月销售"]) / 10
   df201ab8["11月平均销售"] = (df201ab8["1月销售"] + df201ab8["2月销售"] + df201ab8["3月销售"] + df201ab8[
       "4月销售"] +
                          df201ab8["5月销售"] + df201ab8["6月销售"] + df201ab8["7月销售"] + df201ab8["8月销售"] + df201ab8[
                              "9月销售"] + df201ab8["10月销售"] + df201ab8["11月销售"]) / 11
   df201ab8["12月平均销售"] = (df201ab8["1月销售"] + df201ab8["2月销售"] + df201ab8["3月销售"] + df201ab8[
       "4月销售"] +
                          df201ab8["5月销售"] + df201ab8["6月销售"] + df201ab8["7月销售"] + df201ab8["8月销售"] + df201ab8[
                              "9月销售"] + df201ab8["10月销售"] + df201ab8["11月销售"] + df201ab8["12月销售"]) / 12
   df201ab9 = df1400[df1400["业务员"] == "朱津齐"]
   df201ab9["1月平均销售"] = df201ab9["1月销售"]
   df201ab9["2月平均销售"] = (df201ab9["1月销售"] + df201ab9["2月销售"]) / 2
   df201ab9["3月平均销售"] = (df201ab9["1月销售"] + df201ab9["2月销售"] + df201ab9["3月销售"]) / 3
   df201ab9["4月平均销售"] = (df201ab9["1月销售"] + df201ab9["2月销售"] + df201ab9["3月销售"] + df201ab9["4月销售"]) / 4
   df201ab9["5月平均销售"] = (df201ab9["1月销售"] + df201ab9["2月销售"] + df201ab9["3月销售"] + df201ab9[
       "4月销售"] + df201ab9["5月销售"]) / 5
   df201ab9["6月平均销售"] = (df201ab9["1月销售"] + df201ab9["2月销售"] + df201ab9["3月销售"] + df201ab9[
       "4月销售"] +
                         df201ab9["5月销售"] + df201ab9["6月销售"]) / 6
   df201ab9["7月平均销售"] = (df201ab9["1月销售"] + df201ab9["2月销售"] + df201ab9["3月销售"] + df201ab9[
       "4月销售"] +
                         df201ab9["5月销售"] + df201ab9["6月销售"] + df201ab9["7月销售"]) / 7
   df201ab9["8月平均销售"] = (df201ab9["1月销售"] + df201ab9["2月销售"] + df201ab9["3月销售"] + df201ab9[
       "4月销售"] +
                         df201ab9["5月销售"] + df201ab9["6月销售"] + df201ab9["7月销售"] + df201ab9["8月销售"]) / 8
   df201ab9["9月平均销售"] = (df201ab9["1月销售"] + df201ab9["2月销售"] + df201ab9["3月销售"] + df201ab9[
       "4月销售"] +
                         df201ab9["5月销售"] + df201ab9["6月销售"] + df201ab9["7月销售"] + df201ab9["8月销售"] + df201ab9[
                             "9月销售"]) / 9
   df201ab9["10月平均销售"] = (df201ab9["1月销售"] + df201ab9["2月销售"] + df201ab9["3月销售"] + df201ab9[
       "4月销售"] +
                          df201ab9["5月销售"] + df201ab9["6月销售"] + df201ab9["7月销售"] + df201ab9["8月销售"] + df201ab9[
                              "9月销售"] + df201ab9["10月销售"]) / 10
   df201ab9["11月平均销售"] = (df201ab9["1月销售"] + df201ab9["2月销售"] + df201ab9["3月销售"] + df201ab9[
       "4月销售"] +
                          df201ab9["5月销售"] + df201ab9["6月销售"] + df201ab9["7月销售"] + df201ab9["8月销售"] + df201ab9[
                              "9月销售"] + df201ab9["10月销售"] + df201ab9["11月销售"]) / 11
   df201ab9["12月平均销售"] = (df201ab9["1月销售"] + df201ab9["2月销售"] + df201ab9["3月销售"] + df201ab9[
       "4月销售"] +
                          df201ab9["5月销售"] + df201ab9["6月销售"] + df201ab9["7月销售"] + df201ab9["8月销售"] + df201ab9[
                              "9月销售"] + df201ab9["10月销售"] + df201ab9["11月销售"] + df201ab9["12月销售"]) / 12

   df201ab10 = df1400[df1400["业务员"] == "药品2部"]
   df201ab10["1月平均销售"] = df201ab10["1月销售"]
   df201ab10["2月平均销售"] = (df201ab10["1月销售"] + df201ab10["2月销售"]) / 2
   df201ab10["3月平均销售"] = (df201ab10["1月销售"] + df201ab10["2月销售"] + df201ab10["3月销售"]) / 3
   df201ab10["4月平均销售"] = (df201ab10["1月销售"] + df201ab10["2月销售"] + df201ab10["3月销售"] + df201ab10["4月销售"]) / 4
   df201ab10["5月平均销售"] = (df201ab10["1月销售"] + df201ab10["2月销售"] + df201ab10["3月销售"] + df201ab10[
       "4月销售"] + df201ab10["5月销售"]) / 5
   df201ab10["6月平均销售"] = (df201ab10["1月销售"] + df201ab10["2月销售"] + df201ab10["3月销售"] + df201ab10[
       "4月销售"] +
                          df201ab10["5月销售"] + df201ab10["6月销售"]) / 6
   df201ab10["7月平均销售"] = (df201ab10["1月销售"] + df201ab10["2月销售"] + df201ab10["3月销售"] + df201ab10[
       "4月销售"] +
                          df201ab10["5月销售"] + df201ab10["6月销售"] + df201ab10["7月销售"]) / 7
   df201ab10["8月平均销售"] = (df201ab10["1月销售"] + df201ab10["2月销售"] + df201ab10["3月销售"] + df201ab10[
       "4月销售"] +
                          df201ab10["5月销售"] + df201ab10["6月销售"] + df201ab10["7月销售"] + df201ab10["8月销售"]) / 8
   df201ab10["9月平均销售"] = (df201ab10["1月销售"] + df201ab10["2月销售"] + df201ab10["3月销售"] + df201ab10[
       "4月销售"] +
                          df201ab10["5月销售"] + df201ab10["6月销售"] + df201ab10["7月销售"] + df201ab10["8月销售"] + df201ab10[
                              "9月销售"]) / 9
   df201ab10["10月平均销售"] = (df201ab10["1月销售"] + df201ab10["2月销售"] + df201ab10["3月销售"] + df201ab10[
       "4月销售"] +
                           df201ab10["5月销售"] + df201ab10["6月销售"] + df201ab10["7月销售"] + df201ab10["8月销售"] + df201ab10[
                               "9月销售"] + df201ab10["10月销售"]) / 10
   df201ab10["11月平均销售"] = (df201ab10["1月销售"] + df201ab10["2月销售"] + df201ab10["3月销售"] + df201ab10[
       "4月销售"] +
                           df201ab10["5月销售"] + df201ab10["6月销售"] + df201ab10["7月销售"] + df201ab10["8月销售"] + df201ab10[
                               "9月销售"] + df201ab10["10月销售"] + df201ab10["11月销售"]) / 11
   df201ab10["12月平均销售"] = (df201ab10["1月销售"] + df201ab10["2月销售"] + df201ab10["3月销售"] + df201ab10[
       "4月销售"] +
                           df201ab10["5月销售"] + df201ab10["6月销售"] + df201ab10["7月销售"] + df201ab10["8月销售"] + df201ab10[
                               "9月销售"] + df201ab10["10月销售"] + df201ab10["11月销售"] + df201ab10["12月销售"]) / 12

   df1499 = pd.concat([df201ab, df201ab1, df201ab2,df201ab3,df201ab4,df201ab5,df201ab6,df201ab7,df201ab8,df201ab9,df201ab10],
                      ignore_index=True)  # 组合
   df1500 = df1499.drop(
       ["1月销售", "2月销售", "3月销售", "4月销售", "5月销售", "6月销售", "7月销售", "8月销售", "9月销售", "10月销售",
        "11月销售", "12月销售"], axis=1)  # 删列



   ####表5天数计算
   ####应收账款平均余额/销售平均额*30

   df800 = pd.merge(df1300, df1500, how='left', on=['业务员']);
   df800["1月周转天数"] = df800["1月平均余额"]/df800["1月平均销售"]*30
   df800["2月周转天数"] = df800["2月平均余额"] / df800["2月平均销售"] * 30
   df800["3月周转天数"] = df800["3月平均余额"] / df800["3月平均销售"] * 30
   df800["4月周转天数"] = df800["4月平均余额"] / df800["4月平均销售"] * 30
   df800["5月周转天数"] = df800["5月平均余额"] / df800["5月平均销售"] * 30
   df800["6月周转天数"] = df800["6月平均余额"] / df800["6月平均销售"] * 30
   df800["7月周转天数"] = df800["7月平均余额"] / df800["7月平均销售"] * 30
   df800["8月周转天数"] = df800["8月平均余额"] / df800["8月平均销售"] * 30
   df800["9月周转天数"] = df800["9月平均余额"] / df800["9月平均销售"] * 30
   df800["10月周转天数"] = df800["10月平均余额"] / df800["10月平均销售"] * 30
   df800["11月周转天数"] = df800["11月平均余额"] / df800["11月平均销售"] * 30
   df800["12月周转天数"] = df800["12月平均余额"] / df800["12月平均销售"] * 30

   df1598 = df800.drop(
       ["1月平均销售", "2月平均销售", "3月平均销售", "4月平均销售", "5月平均销售", "6月平均销售", "7月平均销售", "8月平均销售", "9月平均销售", "10月平均销售","11月平均销售", "12月平均销售"], axis=1)  # 删列
   df1600 = df1598.drop(
       ["1月平均余额", "2月平均余额", "3月平均余额", "4月平均余额", "5月平均余额", "6月平均余额", "7月平均余额", "8月平均余额", "9月平均余额", "10月平均余额",
        "11月平均余额", "12月平均余额"], axis=1)  # 删列

   df1511 = df1500[df1500["业务员"] != "傅诗云"]
   df1501a = df1511[df1511["业务员"] != "张丽君"]
   df1502a = df1501a[df1501a["业务员"] != "戚肆朝"]
   df1503a = df1502a[df1502a["业务员"] != "王振杰"]
   df1504a = df1503a[df1503a["业务员"] != "徐蕾"]
   df1505a = df1504a[df1504a["业务员"] != "朱津齐"]
   df1506a = df1505a[df1505a["业务员"] != "杨阳"]
   df1507a = df1506a[df1506a["业务员"] != "王伟平"]
   df1508a = df1507a[df1507a["业务员"] != "王宇栋"]

   df1508a.loc['平均销售'] = df1508a.apply(lambda x: x.sum(), axis=0)

   df1508a["业务员"].replace("药品1部药品2部", "合计", inplace=True)
   df1509a = df1508a[df1508a["业务员"] != "药品1部"]
   df1510a = df1509a[df1509a["业务员"] != "药品2部"]

   df1311 = df1300[df1300["业务员"] != "傅诗云"]
   df1312 = df1311[df1311["业务员"] != "张丽君"]
   df1313 = df1312[df1312["业务员"] != "戚肆朝"]
   df1314 = df1313[df1313["业务员"] != "王振杰"]
   df1315 = df1314[df1314["业务员"] != "徐蕾"]
   df1316 = df1315[df1315["业务员"] != "朱津齐"]
   df1317 = df1316[df1316["业务员"] != "杨阳"]
   df1318 = df1317[df1317["业务员"] != "王伟平"]
   df1319 = df1318[df1318["业务员"] != "王宇栋"]

   df1319.loc['平均余额'] = df1319.apply(lambda x: x.sum(), axis=0)

   df1319["业务员"].replace("药品1部药品2部", "合计", inplace=True)
   df1320 = df1319[df1319["业务员"] != "药品1部"]
   df1321 = df1320[df1320["业务员"] != "药品2部"]

   df2000 = pd.merge(df1321, df1510a, how='left', on=['业务员']);
   df2000["1月周转天数"] = df2000["1月平均余额"] / df2000["1月平均销售"] * 30
   df2000["2月周转天数"] = df2000["2月平均余额"] / df2000["2月平均销售"] * 30
   df2000["3月周转天数"] = df2000["3月平均余额"] / df2000["3月平均销售"] * 30
   df2000["4月周转天数"] = df2000["4月平均余额"] / df2000["4月平均销售"] * 30
   df2000["5月周转天数"] = df2000["5月平均余额"] / df2000["5月平均销售"] * 30
   df2000["6月周转天数"] = df2000["6月平均余额"] / df2000["6月平均销售"] * 30
   df2000["7月周转天数"] = df2000["7月平均余额"] / df2000["7月平均销售"] * 30
   df2000["8月周转天数"] = df2000["8月平均余额"] / df2000["8月平均销售"] * 30
   df2000["9月周转天数"] = df2000["9月平均余额"] / df2000["9月平均销售"] * 30
   df2000["10月周转天数"] = df2000["10月平均余额"] / df2000["10月平均销售"] * 30
   df2000["11月周转天数"] = df2000["11月平均余额"] / df2000["11月平均销售"] * 30
   df2000["12月周转天数"] = df2000["12月平均余额"] / df2000["12月平均销售"] * 30

   df2001 = df2000.drop(
       ["1月平均销售", "2月平均销售", "3月平均销售", "4月平均销售", "5月平均销售", "6月平均销售", "7月平均销售", "8月平均销售", "9月平均销售", "10月平均销售", "11月平均销售",
        "12月平均销售"], axis=1)  # 删列
   df2002 = df2001.drop(
       ["1月平均余额", "2月平均余额", "3月平均余额", "4月平均余额", "5月平均余额", "6月平均余额", "7月平均余额", "8月平均余额", "9月平均余额", "10月平均余额",
        "11月平均余额", "12月平均余额"], axis=1)  # 删列

   df2003 = pd.concat([df1600, df2002], ignore_index=True)  # 组合
   #####
   ####df1200表1  余额
   ####df1300表2  平均余额
   ####df1400表3  销售
   ####df1500表4  平均销售
   ####df1600表5  计算天数





   #df1601=df1600
   #df1300.loc['平均余额'] = df1300.apply(lambda x: x.sum(), axis=0)
   #df1600["业务员"].replace("傅诗云戚肆朝张丽君王振杰药品1部徐蕾朱津齐杨阳王伟平王宇栋药品2部", "平均余额", inplace=True)
   #df1602 = df1600[df1600["业务员"] == "平均余额"]

   #df213a = df1200[df1200["业务员"] == "药品2部"]
   #df1605 = pd.concat([df1600, df1604],ignore_index=True)  # 组合




   #df1800.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
    #                                                             filetypes=[("Microsoft Excel文件", "*.xlsx"),
   #                                                                         ("Microsoft Excel 97-20003 文件", "*.xls")],
    #                                                             defaultextension=".xlsx"));


   df2003.to_excel("周期天数" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                   index=False)  # 自动输出
   df1500.to_excel("平均销售" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                   index=False)  # 自动输出
   df1400.to_excel("销售" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                   index=False)  # 自动输出
   df1300.to_excel("平均余额" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                   index=False)  # 自动输出
   df1200.to_excel("余额" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                   index=False)  # 自动输出
   tkinter.messagebox.showinfo("运行结果", "回款天数计算导出成功！");

 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
             message="请检查提交源文件是否正确 '" + str(error) + "'.",
             detail=traceback.format_exc())



def appendStr31():  #############################################################################################客户年度销售整理

    try:
        tkinter.messagebox.showinfo("提醒", "请选择英克导出源文件");
        df1 = pd.read_excel(tkinter.filedialog.askopenfilename());


        df2 = df1.rename(columns={'货品三级分类': '分类代码', '金额': '19年实际英克','基本单位数量':'数量','':''});
        df3 = df2.groupby(["客户ID","客户名称", "通用名","货品id"], as_index=False)["19年实际英克","数量"].sum();

        tkinter.messagebox.showinfo("提醒", "选择客户ID源文件");
        df10 = pd.read_excel(tkinter.filedialog.askopenfilename());
        df4 = pd.merge(df10, df3, how='left', on=['客户ID']);  # 完全相同合并，忽略没有的货品ID(没有how)




        df4.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                        filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                                   (
                                                                                       "Microsoft Excel 97-20003 文件",
                                                                                       "*.xls")],
                                                                        defaultextension=".xlsx"));


        tkinter.messagebox.showinfo("运行结果", "英克导出成功！");
    except Exception as error:
        tm.showerror(title="煎饼提示前方路堵",
                     message="请检查提交源文件是否正确 '" + str(error) + "'.",
                     detail=traceback.format_exc())



def appendStr32():  #######################################################################################################20200129新分类考核表
    try:

        tkinter.messagebox.showinfo("提醒", "请先选择金蝶科目核算项目文件");

        df1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径
        # df2 =df1.drop([2], axis=1)  # 删列

        df2 = df1['项目代码'].str.split('-| ', expand=True);
        df3 = pd.merge(df2, df1, right_index=True, left_index=True);  # df3表为科目表总表原表

        # df3.to_excel(excel_writer="D:/合并测试/合并测试数据2.xlsx",
        #       sheet_name="处理",
        #       index=False);

        df3["部门"] = df3[0]
        df3["片区"] = df3[1]
        df3["片区编码"] = df3[2]
        df3["本月试剂销售"] = df3["本期发生额"]
        df3["本月试剂收款"] = df3['Unnamed: 10']

        df31a = df3
        # df31a.to_excel(excel_writer="D:/合并测试数据3.xlsx",
        #        sheet_name="处理",
        #       );

        # 以上是复制列名到表最后

        df5 = df3[df3["项目"] != "仪器"]
        df51 = df5[df5["项目"] != "代理药品"]
        df52 = df51[df51["公司"] != "海尔施集团"]
        # 去除仪器类,剩余非仪器类

        # df82.groupby(["客户名称", "业务员"], as_index=False)["过期", "Unnamed: 6", "Unnamed: 7", "Unnamed: 8", "Unnamed: 9"].sum();

        df6 = df52.groupby(["部门", "片区", "片区编码", "公司"])[
            "本月试剂销售", "本月试剂收款", "期初余额", "期末余额", "Unnamed: 8", "Unnamed: 14"].sum();
        c_df = pd.DataFrame(df6)
        c_df.reset_index(inplace=True)  # 取消合并
        df6["期初应收余额"] = df6["期初余额"] - df6["Unnamed: 8"]
        df6["期末应收余额"] = df6["期末余额"] - df6["Unnamed: 14"]

        df7 = df6.drop(["Unnamed: 8", "Unnamed: 14", "期初余额", "期末余额"], axis=1)  # 删列
        df8 = df7.sort_values(by=['片区编码'], axis=0, ascending=True)  # 行排序

       # df8.to_excel(excel_writer="D:/2020新表/测试1311.xlsx",
        #             sheet_name="处理",
         #           );

        # df81 = pd.pivot_table(df8, index=["部门","片区","片区编码"], columns="公司", values=["本月试剂销售","本月试剂收款","期初应收余额","期末应收余额"]);  # 列换行公司排列#######################有用0708
        # c_df = pd.DataFrame(df81)
        # c_df.reset_index(inplace=True)

        # df81.to_excel(excel_writer="D:/合并测试/合并测试数据3.xlsx",
        #             sheet_name="处理",
        #            );

        print(df8)
        #####以上试剂销售收款完毕

        ######加入表头防止没有发生遗漏

        ###开始仪器销售收款

        df9 = df31a[df31a["项目"] != "代理试剂"]
        df10 = df9[df9["项目"] != "配件"]
        df11 = df10[df10["项目"] != "代理药品"]
        df12 = df11[df11["项目"] != "其他（含服务费）"]
        df13 = df12[df12["项目"] != "维修"]
        df14 = df13[df13["项目"] != "仪器租赁"]
        df141 = df14[df14["项目"] != "代理药品"]
        #生物简易征收
        df14102 = df141[df141["项目"] != "生物简易征收"]
        df1411 = df14102[df14102["项目"] != "自营检验服务"]  # 8.8修改检验所检验服务不应该划为仪器类

        df142 = df1411[df1411["公司"] != "海尔施集团"]
        df143 = df142[df142["客户"] != "合计"]

        # df143.to_excel(excel_writer="D:/合并测试/合并测试数据2.xlsx",
        #    sheet_name="处理",
        #    index=False);

        # df13["本月仪器销售"] = df13["本月试剂销售"]
        # df13["本月仪器收款"] = df13["本月试剂收款"]

        df15 = df143.groupby(["部门", "片区", "片区编码", "公司"], as_index=False)["Unnamed: 10", "本期发生额"].sum();

        df15["本期仪器收款"] = df15['Unnamed: 10']
        df15["本期仪器销售"] = df15["本期发生额"]

        df16 = df15.drop(["Unnamed: 10", "本期发生额"], axis=1)  # 删列

        # df16.to_excel(excel_writer="D:/合并测试/合并测试数据0807.xlsx",
        #        sheet_name="处理",
        #      );

        ########以上仪器本月收款销售完毕

        df17 = pd.merge(df8, df16, how='outer', on=['片区', '部门', '公司', '片区编码']);  # 完全相同合并，忽略没有的客户(没有how)

        #df17.to_excel(excel_writer="D:/2020新表/08072.xlsx",
         #             sheet_name="处理",
          #           index=False);

        # df18 = df17.drop(["片区编码_y"], axis=1)  # 删列
        df181 = df17[df17["部门"] != ""]  # 去除部门为空的记录
        df19 = df181.sort_values(by=['片区编码'], axis=0, ascending=True)  # 行排序

        df19["片区编码"] = df19["片区"]

        df19["片区编码"].replace("温州1", "0101", inplace=True)
        df19["片区编码"].replace("温州2", "0102", inplace=True)
        df19["片区编码"].replace("台州1", "0103", inplace=True)
        df19["片区编码"].replace("台州2", "0104", inplace=True)
        df19["片区编码"].replace("丽水", "0105", inplace=True)
        df19["片区编码"].replace("宁波", "0201", inplace=True)
        df19["片区编码"].replace("舟山北仑", "0202", inplace=True)
        df19["片区编码"].replace("北三县", "0203", inplace=True)
        df19["片区编码"].replace("南三县", "0204", inplace=True)
        df19["片区编码"].replace("杭州省级", "0301", inplace=True)
        df19["片区编码"].replace("杭州市级", "0302", inplace=True)  #
        df19["片区编码"].replace("嘉湖", "0303", inplace=True)  #
        df19["片区编码"].replace("南京吴珏", "0401", inplace=True)   #0129修改
        df19["片区编码"].replace("南京高跃", "0402", inplace=True)
        df19["片区编码"].replace("南通李国旺", "0501", inplace=True)
        df19["片区编码"].replace("盐城", "0503", inplace=True)
        df19["片区编码"].replace("连云港", "0504", inplace=True)
        df19["片区编码"].replace("上海1", "0601", inplace=True)
        df19["片区编码"].replace("上海2", "0602", inplace=True)
        df19["片区编码"].replace("上海3", "0603", inplace=True)    #0129修改
        df19["片区编码"].replace("苏州", "0701", inplace=True)
        df19["片区编码"].replace("苏州市郊", "0702", inplace=True)
        df19["片区编码"].replace("扬州", "0801", inplace=True)
        df19["片区编码"].replace("泰州", "0802", inplace=True)
        df19["片区编码"].replace("徐州", "0901", inplace=True)  #
        df19["片区编码"].replace("宿迁", "0902", inplace=True)
        df19["片区编码"].replace("淮安", "0903", inplace=True)
        df19["片区编码"].replace("常州", "1001", inplace=True)
        df19["片区编码"].replace("镇江", "1002", inplace=True)
        df19["片区编码"].replace("绍兴龚群波", "1101", inplace=True)
        df19["片区编码"].replace("金衢", "1102", inplace=True)  #
        df19["片区编码"].replace("无锡1", "1201", inplace=True)
        df19["片区编码"].replace("无锡2", "1202", inplace=True)
        df19["片区编码"].replace("业务拓展部", "1301", inplace=True)
        df19["片区编码"].replace("调拨", "1901", inplace=True)
        df19["片区编码"].replace("配件", "1902", inplace=True)
        df19["片区编码"].replace("生化分销", "1903", inplace=True)


        # df191 = df19.drop(["片区编码"], axis=1)  # 删列

        df191 = df19
        #df191.to_excel(excel_writer="D:/2020新表/2020新表测试.xlsx",
         #   sheet_name="处理",
          #  index=False);

        ####导入集团账龄
        tkinter.messagebox.showinfo("提醒", "请选择集团账龄源文件");
        df100 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径
        df101 = df100['客户'].str.split('-| ', expand=True);
        df102 = pd.merge(df101, df100, right_index=True, left_index=True);

        df102["部门"] = df102[0]
        df102["片区"] = df102[1]
        df102["片区编码_x"] = df102[2]

        # df102.to_excel(excel_writer="D:/合并测试/合并测试数据2.xlsx",
        #      sheet_name="处理",
        #     index=False);

        df103 = df102[df102["项目"] != "003 仪器"]
        df104 = df103[df103["项目"] != "007 代理药品"]

        df105 = df104.groupby(["部门", "片区", "片区编码_x", "公司"], as_index=False)["Unnamed: 10"].sum();

        df106 = df105[df105["部门"] != ""]

        df106["片区编码"] = df106["片区"]

        df106["片区编码"].replace("温州1", "0101", inplace=True)
        df106["片区编码"].replace("温州2", "0102", inplace=True)
        df106["片区编码"].replace("台州1", "0103", inplace=True)
        df106["片区编码"].replace("台州2", "0104", inplace=True)
        df106["片区编码"].replace("丽水", "0105", inplace=True)
        df106["片区编码"].replace("宁波", "0201", inplace=True)
        df106["片区编码"].replace("舟山北仑", "0202", inplace=True)
        df106["片区编码"].replace("北三县", "0203", inplace=True)
        df106["片区编码"].replace("南三县", "0204", inplace=True)
        df106["片区编码"].replace("杭州省级", "0301", inplace=True)
        df106["片区编码"].replace("杭州市级", "0302", inplace=True)  #
        df106["片区编码"].replace("嘉湖", "0303", inplace=True)  #
        df106["片区编码"].replace("南京吴珏", "0401", inplace=True)  # 0129修改
        df106["片区编码"].replace("南京高跃", "0402", inplace=True)
        df106["片区编码"].replace("南通李国旺", "0501", inplace=True)
        df106["片区编码"].replace("盐城", "0503", inplace=True)
        df106["片区编码"].replace("连云港", "0504", inplace=True)
        df106["片区编码"].replace("上海1", "0601", inplace=True)
        df106["片区编码"].replace("上海2", "0602", inplace=True)
        df106["片区编码"].replace("上海3", "0603", inplace=True)  # 0129修改
        df106["片区编码"].replace("苏州", "0701", inplace=True)
        df106["片区编码"].replace("苏州市郊", "0702", inplace=True)
        df106["片区编码"].replace("扬州", "0801", inplace=True)
        df106["片区编码"].replace("泰州", "0802", inplace=True)
        df106["片区编码"].replace("徐州", "0901", inplace=True)  #
        df106["片区编码"].replace("宿迁", "0902", inplace=True)
        df106["片区编码"].replace("淮安", "0903", inplace=True)
        df106["片区编码"].replace("常州", "1001", inplace=True)
        df106["片区编码"].replace("镇江", "1002", inplace=True)
        df106["片区编码"].replace("绍兴龚群波", "1101", inplace=True)
        df106["片区编码"].replace("金衢", "1102", inplace=True)  #
        df106["片区编码"].replace("无锡1", "1201", inplace=True)
        df106["片区编码"].replace("无锡2", "1202", inplace=True)
        df106["片区编码"].replace("业务拓展部", "1301", inplace=True)
        df106["片区编码"].replace("调拨", "1901", inplace=True)
        df106["片区编码"].replace("配件", "1902", inplace=True)
        df106["片区编码"].replace("生化分销", "1903", inplace=True)


        df107 = df106.drop(["片区编码_x", "片区"], axis=1)  # 删列

        #df107.to_excel(excel_writer="D:/2020新表/合并测试数据2.xlsx",
         #      sheet_name="处理",
          #    index=False);

        ####俩表新建部门字段，修改为编码，然后合并0707

        ####拼接接团账龄
        # df107 = pd.merge(df19,df106, right_index=True, left_index=True);

        df107 = pd.merge(df191, df107, how='left', on=['片区编码', '部门', '公司']);  # 完全相同合并
        df107["三个月以上应收账款"] = df107["Unnamed: 10"]

        df1071 = df107.drop(["Unnamed: 10"], axis=1)  # 删列 df108 改1071

        dfa108 = df1071[df1071["片区"] == "调拨检验所"]
        dfa108["部门"].replace("调拨", "调拨-检验所", inplace=True)

        dfa110a = df1071[df1071["片区"] != "调拨检验所"]
        df108 = pd.concat([dfa110a, dfa108], ignore_index=True)
        #df108.to_excel(excel_writer="D:/2020新表/二稿测试11294.xlsx",
        #            sheet_name="处理",
         #          index=False);
        # dfa110a.to_excel(excel_writer="D:/合并测试/二稿测试11295.xlsx",
        #                sheet_name="处理",
        #                index=False);
        #######以上是试剂和仪器收款拼接完毕,下面开始增加部门小计（调拨类的分开统计）
        dfa200 = df108[df108["部门"] != "调拨-检验所"]
        df201 = dfa200[dfa200["部门"] != "调拨"]  # 去除部门为调拨，配件，生化分销
        df201 = df201[df201["部门"] != "配件"]  # 去除部门为调拨，配件，生化分销
        df202 = df201[df201["部门"] != "生化分销"]  #

        #0203去除部门中的检验所后合计
        df203 = df202[df202["片区"] != "一部检验所"]
        df204 = df203[df203["片区"] != "二部检验所"]
        df205 = df204[df204["片区"] != "三部检验所"]
        df206 = df205[df205["片区"] != "四部检验所"]
        df207 = df206[df206["片区"] != "五部检验所"]
        df208 = df207[df207["片区"] != "六部检验所"]
        df209 = df208[df208["片区"] != "七部检验所"]
        df210 = df209[df209["片区"] != "八部检验所"]
        df211 = df210[df210["片区"] != "九部检验所"]
        df212 = df211[df211["片区"] != "十部检验所"]
        df213 = df212[df212["片区"] != "十一部检验所"]
        df214 = df213[df213["片区"] != "十二部检验所"]



        df20 = df214.groupby(["部门"], as_index=False)[
            "本月试剂销售", "本月试剂收款", "期初应收余额", "期末应收余额", "本期仪器收款", "本期仪器销售", "三个月以上应收账款"].sum();
        df20["片区"] = df20["部门"]
        df20["片区编码"] = df20["部门"]
        df20["公司"] = df20["部门"]

        df20["片区"].replace("一部", "诊断一部", inplace=True)
        df20["片区"].replace("二部", "诊断二部", inplace=True)
        df20["片区"].replace("三部", "诊断三部", inplace=True)
        df20["片区"].replace("四1部", "诊断四1部", inplace=True)
        df20["片区"].replace("四2部", "诊断四2部", inplace=True)
        df20["片区"].replace("五1部", "诊断五1部", inplace=True)
        df20["片区"].replace("五2部", "诊断五2部", inplace=True)
        df20["片区"].replace("六部", "诊断六部", inplace=True)
        df20["片区"].replace("七部", "诊断七部", inplace=True)
        df20["片区"].replace("八部", "诊断八部", inplace=True)
        df20["片区"].replace("九部", "诊断九部", inplace=True)
        df20["片区"].replace("十部", "诊断十部", inplace=True)
        df20["片区"].replace("十一部", "诊断十一部", inplace=True)
        df20["片区"].replace("十二部", "诊断十二部", inplace=True)
        df20["片区"].replace("业务拓展部", "业务拓展部", inplace=True)   #0129修改


        df20["片区编码"].replace("一部", "0199", inplace=True)
        df20["片区编码"].replace("二部", "0299", inplace=True)
        df20["片区编码"].replace("三部", "0399", inplace=True)
        df20["片区编码"].replace("四1部", "04019", inplace=True)
        df20["片区编码"].replace("四2部", "04029", inplace=True)
        df20["片区编码"].replace("五1部", "05019", inplace=True)
        df20["片区编码"].replace("五2部", "05049", inplace=True)
        df20["片区编码"].replace("六部", "0699", inplace=True)
        df20["片区编码"].replace("七部", "0799", inplace=True)
        df20["片区编码"].replace("八部", "0899", inplace=True)
        df20["片区编码"].replace("九部", "0999", inplace=True)
        df20["片区编码"].replace("十部", "1099", inplace=True)
        df20["片区编码"].replace("十一部", "1199", inplace=True)
        df20["片区编码"].replace("十二部", "1299", inplace=True)
        df20["片区编码"].replace("业务拓展部", "1399", inplace=True)




        df20["公司"].replace("一部", "总计", inplace=True)
        df20["公司"].replace("二部", "总计", inplace=True)
        df20["公司"].replace("三部", "总计", inplace=True)
        df20["公司"].replace("四1部", "总计", inplace=True)
        df20["公司"].replace("四2部", "总计", inplace=True)
        df20["公司"].replace("五1部", "总计", inplace=True)
        df20["公司"].replace("五2部", "总计", inplace=True)
        df20["公司"].replace("六部", "总计", inplace=True)
        df20["公司"].replace("七部", "总计", inplace=True)
        df20["公司"].replace("八部", "总计", inplace=True)
        df20["公司"].replace("九部", "总计", inplace=True)
        df20["公司"].replace("十部", "总计", inplace=True)
        df20["公司"].replace("十一部", "总计", inplace=True)
        df20["公司"].replace("十二部", "总计", inplace=True)
        df20["公司"].replace("业务拓展部", "总计", inplace=True)
        df21 = pd.concat([df108, df20], ignore_index=True)
        df211 = df21[df21["部门"] != ""]

        df22 = df211.sort_values(by=['片区编码'], axis=0, ascending=True)  # 行排序

        ####初步组合完毕

        ####剥离调拨部门

        df23 = df22[df22["部门"] != "一部"]
        df231 = df23[df23["部门"] != "二部"]
        df232 = df231[df231["部门"] != "三部"]
        df233 = df232[df232["部门"] != "四1部"]
        df234 = df233[df233["部门"] != "五1部"]
        df235 = df234[df234["部门"] != "六部"]
        df236 = df235[df235["部门"] != "七部"]
        df237 = df236[df236["部门"] != "八部"]
        df238 = df237[df237["部门"] != "九部"]
        df239 = df238[df238["部门"] != "十部"]
        df240 = df239[df239["部门"] != "十一部"]
        df241 = df240[df240["部门"] != "十二部"]
        df242 = df241[df241["部门"] != "业务拓展部"]
        df243 = df242[df242["部门"] != "四2部"]
        df244 = df243[df243["部门"] != "五2部"]
        df2411 = df244[df244["片区"] != "调拨检验所"]
        df2411["片区编码"] = df2411["部门"]

        # df241.to_excel(excel_writer="D:/合并测试/二稿测试1.xlsx",sheet_name="处理",index=False);

        df2411["片区编码"].replace("调拨", "1900", inplace=True)
        df2411["片区编码"].replace("配件", "1900", inplace=True)
        df2411["片区编码"].replace("生化分销", "1900", inplace=True)

        df24 = df2411.groupby(["片区编码", "公司"], as_index=False)[
            "本月试剂销售", "本月试剂收款", "期初应收余额", "期末应收余额", "本期仪器收款", "本期仪器销售", "三个月以上应收账款"].sum();

        df24["部门"] = df24["片区编码"]
        df24["片区"] = df24["片区编码"]
        # df24["公司"] = df24["片区编码"]

        df24["部门"].replace("1900", "调拨", inplace=True)
        df24["片区"].replace("1900", "调拨部", inplace=True)
        # df24["公司"].replace("1900", "总计", inplace=True)

        #df24.to_excel(excel_writer="D:/2020新表/调拨合计测试.xlsx",
         #             sheet_name="处理",
          #           index=False);

        df25 = df22[df22["部门"] != "调拨"]
        df26 = df25[df25["部门"] != "配件"]
        df27 = df26[df26["部门"] != "生化分销"]

        df28 = pd.concat([df27, df24], ignore_index=True)

        #0203新增调拨合计

        df2811=df24.groupby(["片区编码"], as_index=False)[
            "本月试剂销售", "本月试剂收款", "期初应收余额", "期末应收余额", "本期仪器收款", "本期仪器销售", "三个月以上应收账款"].sum();
        df2811["公司"]=df2811["片区编码"]
        df2811["片区"]=df2811["片区编码"]
        df2811["公司"].replace("1900", "总计", inplace=True)
        df2811["片区编码"].replace("1900", "1999", inplace=True)
        df2811["片区"].replace("1900", "调拨部", inplace=True)

        df281 = pd.concat([df28, df2811], ignore_index=True)



        #####调拨整合完毕,引用模板为df28修改三部,九部,十一部,汇总地区


        df63 = df281.sort_values(by=['片区编码'], axis=0, ascending=True)  # 行排序
        print(df63)


        #####数据组合完毕,开始表格列排序

        df63["部门名称"] = df63["部门"]
        df63["地区"] = df63["片区"]
        df63["负责人"] = df63["片区"]
        df63["地区编码"] = df63["片区编码"]
        df63["本期试剂销售"] = df63["本月试剂销售"]
        df63["本期试剂收款"] = df63["本月试剂收款"]
        df63["期初余额"] = df63["期初应收余额"]
        df63["期末余额"] = df63["期末应收余额"]
        df63["本月仪器收款"] = df63["本期仪器收款"]
        df63["本月仪器销售"] = df63["本期仪器销售"]
        df63["超三个月应收账款"] = df63["三个月以上应收账款"]

        df64 = df63.drop(["部门", "片区", "片区编码", "本月试剂销售", "本月试剂收款", "期初应收余额", "期末应收余额", "本期仪器收款", "本期仪器销售", "三个月以上应收账款"],
                         axis=1)  # 删列




        #df64.to_excel(excel_writer="D:/2020新表/131.xlsx",  #####710
        #              sheet_name="处理",
        #              );
        ####开始修改片区负责人
        df64["负责人"].replace("温州1", "葛瑞", inplace=True)
        df64["负责人"].replace("温州2", "潘磊", inplace=True)
        df64["负责人"].replace("台州1", "唐惠", inplace=True)
        df64["负责人"].replace("台州2", "胡文魁", inplace=True)
        df64["负责人"].replace("丽水", "方汝泼", inplace=True)
        df64["负责人"].replace("诊断一部", "郭德春", inplace=True)

        df64["负责人"].replace("宁波", "丁玲", inplace=True)
        df64["负责人"].replace("舟山北仑", "高大勇", inplace=True)
        df64["负责人"].replace("北三县", "陆金耀", inplace=True)
        df64["负责人"].replace("南三县", "吴燕江", inplace=True)
        df64["负责人"].replace("诊断二部", "余顶峰", inplace=True)

        df64["负责人"].replace("杭州省级", "姜立民", inplace=True)
        df64["负责人"].replace("杭州市级", "沈剑芳", inplace=True)  #
        df64["负责人"].replace("嘉湖", "阮芳", inplace=True)  #
        df64["负责人"].replace("诊断三部", "毛存亮", inplace=True)

        df64["负责人"].replace("南京高跃", "高跃", inplace=True)
        df64["负责人"].replace("南京吴珏", "吴珏", inplace=True)
        df64["负责人"].replace("诊断四2部", "高跃", inplace=True)
        df64["负责人"].replace("诊断四1部", "吴珏", inplace=True)

        df64["负责人"].replace("南通李国旺", "李国旺", inplace=True)
        df64["负责人"].replace("盐城", "岑潭泽 潘前进", inplace=True)
        df64["负责人"].replace("连云港", "胡士艳 姜健", inplace=True)
        df64["负责人"].replace("诊断五1部", "李国旺", inplace=True)
        df64["负责人"].replace("诊断五2部", "胡士艳", inplace=True)

        df64["负责人"].replace("上海1", "邬幼波", inplace=True)
        df64["负责人"].replace("上海2", "汤俊", inplace=True)
        df64["负责人"].replace("上海3", "黄帅", inplace=True)
        df64["负责人"].replace("诊断六部", "毛存亮", inplace=True)

        df64["负责人"].replace("苏州", "陈凯", inplace=True)
        df64["负责人"].replace("苏州市郊", "吕楠", inplace=True)
        df64["地区"].replace("苏州", "苏州市区", inplace=True)
        df64["地区"].replace("苏州市郊", "苏州郊县", inplace=True)
        df64["负责人"].replace("诊断七部", "全英娜", inplace=True)

        df64["负责人"].replace("扬州", "邹海洵", inplace=True)
        df64["负责人"].replace("泰州", "金英明", inplace=True)
        df64["负责人"].replace("诊断八部", "金英明 邹海洵", inplace=True)

        df64["负责人"].replace("徐州", "唐维洲 于博", inplace=True)  #

        df64["负责人"].replace("宿迁", "赵晨阳 王涛", inplace=True)
        df64["负责人"].replace("淮安", "赵晨阳 白虹", inplace=True)
        df64["负责人"].replace("诊断九部", "吴蓓", inplace=True)

        df64["负责人"].replace("常州", "梅晓虹", inplace=True)
        df64["负责人"].replace("镇江", "于亚惠", inplace=True)
        df64["负责人"].replace("诊断十部", "梅晓虹", inplace=True)

        df64["地区"].replace("绍兴龚群波", "绍兴", inplace=True)
        df64["负责人"].replace("绍兴龚群波", "龚群波", inplace=True)

        df64["负责人"].replace("金衢", "胡迪锋", inplace=True)  #
        df64["负责人"].replace("诊断十一部", "李征", inplace=True)

        df64["负责人"].replace("无锡1", "裘涌", inplace=True)
        df64["负责人"].replace("无锡2", "张立伟 赵飞", inplace=True)
        df64["负责人"].replace("诊断十二部", "邬幼波", inplace=True)

        df64["负责人"].replace("调拨部", "孙婷婷", inplace=True)
        df64["负责人"].replace("业务拓展部", "叶仲华", inplace=True)

        df64["地区编码"].replace("一部检验所", "5001", inplace=True)
        df64["地区编码"].replace("二部检验所", "5002", inplace=True)
        df64["地区编码"].replace("三部检验所", "5003", inplace=True)
        df64["地区编码"].replace("四部检验所", "5004", inplace=True)
        df64["地区编码"].replace("五部检验所", "5005", inplace=True)
        df64["地区编码"].replace("六部检验所", "5006", inplace=True)
        df64["地区编码"].replace("七部检验所", "5007", inplace=True)
        df64["地区编码"].replace("八部检验所", "5008", inplace=True)
        df64["地区编码"].replace("九部检验所", "5009", inplace=True)
        df64["地区编码"].replace("十部检验所", "5010", inplace=True)
        df64["地区编码"].replace("十一部检验所", "5011", inplace=True)
        df64["地区编码"].replace("十二部检验所", "5012", inplace=True)
        df64["地区编码"].replace("调拨检验所", "5013", inplace=True)

        df65 = df64.sort_values(by=['地区编码'], axis=0, ascending=True)  # 行排序


        df66 = df65[df65["部门名称"] != "诊断二部"]
        df67 = df66[df66["部门名称"] != "诊断一部"]

        #0203去掉一部检验所等显示
        df68 = df67[df67["地区"] != "一部检验所"]
        df69 = df68[df68["地区"] != "二部检验所"]
        df70 = df69[df69["地区"] != "三部检验所"]
        df71 = df70[df70["地区"] != "四部检验所"]
        df72 = df71[df71["地区"] != "五部检验所"]
        df73 = df72[df72["地区"] != "六部检验所"]
        df74 = df73[df73["地区"] != "七部检验所"]
        df75 = df74[df74["地区"] != "八部检验所"]
        df76 = df75[df75["地区"] != "九部检验所"]
        df77 = df76[df76["地区"] != "十部检验所"]
        df78 = df77[df77["地区"] != "十一部检验所"]
        df79 = df78[df78["地区"] != "十二部检验所"]
        df80 = df79[df79["地区"] != "调拨检验所"]

        #全表合计 df82
        df81 = df80[df80["公司"] == "总计"]
        df82 = df81.groupby(["公司"], as_index=False)[
            "本期试剂销售", "本期试剂收款", "期初余额", "期末余额", "本月仪器收款", "本月仪器销售", "超三个月应收账款"].sum();

        df82["公司"].replace("总计", "合计", inplace=True)

        df82["地区编码"]=df82["公司"]
        df82["地区"] = df82["公司"]
        df82["地区编码"].replace("总计", "9999", inplace=True)
        df82["地区"].replace("总计", "合计", inplace=True)

        df83 = pd.concat([df80, df82], ignore_index=True)

        #全表排序
        df84 = df83.groupby(["公司","地区","地区编码"], as_index=False)[
            "本期试剂销售", "本期试剂收款", "期初余额", "期末余额", "本月仪器收款", "本月仪器销售", "超三个月应收账款"].sum();

        df85 = df84.sort_values(by=['地区编码'], axis=0, ascending=True)  # 行排序

        df85.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                        filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                                   (
                                                                                       "Microsoft Excel 97-20003 文件",
                                                                                       "*.xls")],
                                                                        defaultextension=".xlsx"));

        tkinter.messagebox.showinfo("运行结果", "考核表2020汇总版导出成功！");


        #df10065.to_excel(excel_writer="D:/2020新表/0807.xlsx",#####710
         #            sheet_name="处理",
          #                         );

        ##################################第二稿结束

        ###################################################################################################################2020-02-04 新需求 加客户明细
        df6 = df52.groupby(["部门", "片区", "片区编码", "公司","客户"])[
            "本月试剂销售", "本月试剂收款", "期初余额", "期末余额", "Unnamed: 8", "Unnamed: 14"].sum();
        c_df = pd.DataFrame(df6)
        c_df.reset_index(inplace=True)  # 取消合并
        df6["期初应收余额"] = df6["期初余额"] - df6["Unnamed: 8"]
        df6["期末应收余额"] = df6["期末余额"] - df6["Unnamed: 14"]

        df7 = df6.drop(["Unnamed: 8", "Unnamed: 14", "期初余额", "期末余额"], axis=1)  # 删列
        df8 = df7.sort_values(by=['片区编码'], axis=0, ascending=True)  # 行排序

        # df8.to_excel(excel_writer="D:/测试72.xlsx",
        #              sheet_name="处理",
        #            );

        # df81 = pd.pivot_table(df8, index=["部门","片区","片区编码"], columns="公司", values=["本月试剂销售","本月试剂收款","期初应收余额","期末应收余额"]);  # 列换行公司排列#######################有用0708
        # c_df = pd.DataFrame(df81)
        # c_df.reset_index(inplace=True)

        # df81.to_excel(excel_writer="D:/合并测试/合并测试数据3.xlsx",
        #             sheet_name="处理",
        #            );

        print(df8)
        #####以上试剂销售收款完毕

        ######加入表头防止没有发生遗漏

        ###开始仪器销售收款

        df9 = df31a[df31a["项目"] != "代理试剂"]
        df10 = df9[df9["项目"] != "配件"]
        df11 = df10[df10["项目"] != "代理药品"]
        df12 = df11[df11["项目"] != "其他（含服务费）"]
        df13 = df12[df12["项目"] != "维修"]
        df14 = df13[df13["项目"] != "仪器租赁"]
        df141 = df14[df14["项目"] != "代理药品"]

        df14102 = df141[df141["项目"] != "生物简易征收"]
        df1411 = df14102[df14102["项目"] != "自营检验服务"]  # 8.8修改检验所检验服务不应该划为仪器类

        df142 = df1411[df1411["公司"] != "海尔施集团"]
        df143 = df142[df142["客户"] != "合计"]

        # df143.to_excel(excel_writer="D:/合并测试/合并测试数据2.xlsx",
        #    sheet_name="处理",
        #    index=False);

        # df13["本月仪器销售"] = df13["本月试剂销售"]
        # df13["本月仪器收款"] = df13["本月试剂收款"]

        df15 = df143.groupby(["部门", "片区", "片区编码", "公司","客户"], as_index=False)["Unnamed: 10", "本期发生额"].sum();

        df15["本期仪器收款"] = df15['Unnamed: 10']
        df15["本期仪器销售"] = df15["本期发生额"]

        df16 = df15.drop(["Unnamed: 10", "本期发生额"], axis=1)  # 删列

        # df16.to_excel(excel_writer="D:/合并测试/合并测试数据0807.xlsx",
        #        sheet_name="处理",
        #      );

        ########以上仪器本月收款销售完毕

        df17 = pd.merge(df8, df16, how='outer', on=['片区', '部门', '公司', '片区编码','客户']);  # 完全相同合并，忽略没有的客户(没有how)

        # df17.to_excel(excel_writer="D:/2020新表/08072.xlsx",
        #             sheet_name="处理",
        #           index=False);

        # df18 = df17.drop(["片区编码_y"], axis=1)  # 删列
        df181 = df17[df17["部门"] != ""]  # 去除部门为空的记录
        df19 = df181.sort_values(by=['片区编码'], axis=0, ascending=True)  # 行排序

        df19["片区编码"] = df19["片区"]

        df19["片区编码"].replace("温州1", "0101", inplace=True)
        df19["片区编码"].replace("温州2", "0102", inplace=True)
        df19["片区编码"].replace("台州1", "0103", inplace=True)
        df19["片区编码"].replace("台州2", "0104", inplace=True)
        df19["片区编码"].replace("丽水", "0105", inplace=True)
        df19["片区编码"].replace("宁波", "0201", inplace=True)
        df19["片区编码"].replace("舟山北仑", "0202", inplace=True)
        df19["片区编码"].replace("北三县", "0203", inplace=True)
        df19["片区编码"].replace("南三县", "0204", inplace=True)
        df19["片区编码"].replace("杭州省级", "0301", inplace=True)
        df19["片区编码"].replace("杭州市级", "0302", inplace=True)  #
        df19["片区编码"].replace("嘉湖", "0303", inplace=True)  #
        df19["片区编码"].replace("南京吴珏", "0401", inplace=True)  # 0129修改
        df19["片区编码"].replace("南京高跃", "0402", inplace=True)
        df19["片区编码"].replace("南通李国旺", "0501", inplace=True)
        df19["片区编码"].replace("盐城", "0503", inplace=True)
        df19["片区编码"].replace("连云港", "0504", inplace=True)
        df19["片区编码"].replace("上海1", "0601", inplace=True)
        df19["片区编码"].replace("上海2", "0602", inplace=True)
        df19["片区编码"].replace("上海3", "0603", inplace=True)  # 0129修改
        df19["片区编码"].replace("苏州", "0701", inplace=True)
        df19["片区编码"].replace("苏州市郊", "0702", inplace=True)
        df19["片区编码"].replace("扬州", "0801", inplace=True)
        df19["片区编码"].replace("泰州", "0802", inplace=True)
        df19["片区编码"].replace("徐州", "0901", inplace=True)  #
        df19["片区编码"].replace("宿迁", "0902", inplace=True)
        df19["片区编码"].replace("淮安", "0903", inplace=True)
        df19["片区编码"].replace("常州", "1001", inplace=True)
        df19["片区编码"].replace("镇江", "1002", inplace=True)
        df19["片区编码"].replace("绍兴龚群波", "1101", inplace=True)
        df19["片区编码"].replace("金衢", "1102", inplace=True)  #
        df19["片区编码"].replace("无锡1", "1201", inplace=True)
        df19["片区编码"].replace("无锡2", "1202", inplace=True)
        df19["片区编码"].replace("业务拓展部", "1301", inplace=True)
        df19["片区编码"].replace("调拨", "1901", inplace=True)
        df19["片区编码"].replace("配件", "1902", inplace=True)
        df19["片区编码"].replace("生化分销", "1903", inplace=True)

        # df191 = df19.drop(["片区编码"], axis=1)  # 删列

        df191 = df19
        # df191.to_excel(excel_writer="D:/2020新表/2020新表测试.xlsx",
        #   sheet_name="处理",
        #  index=False);

        ####导入集团账龄
       # tkinter.messagebox.showinfo("提醒", "请选择集团账龄源文件");
       # df100 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630自主选择路径
        df101 = df100['客户'].str.split('-| ', expand=True);
        df102 = pd.merge(df101, df100, right_index=True, left_index=True);

        df102["部门"] = df102[0]
        df102["片区"] = df102[1]
        df102["片区编码_x"] = df102[2]
        df1021 = df102.rename(columns={'客户': '编码'});
        df1021["客户"] = df1021[7]


        # df1021.to_excel(excel_writer="D:/100.xlsx",
        #      sheet_name="处理",
        #     index=False);

        df103 = df1021[df1021["项目"] != "003 仪器"]
        df104 = df103[df103["项目"] != "007 代理药品"]

        df105 = df104.groupby(["部门", "片区", "片区编码_x", "公司","客户"], as_index=False)["Unnamed: 10"].sum();

        df106 = df105[df105["部门"] != ""]

        df106["片区编码"] = df106["片区"]

        df106["片区编码"].replace("温州1", "0101", inplace=True)
        df106["片区编码"].replace("温州2", "0102", inplace=True)
        df106["片区编码"].replace("台州1", "0103", inplace=True)
        df106["片区编码"].replace("台州2", "0104", inplace=True)
        df106["片区编码"].replace("丽水", "0105", inplace=True)
        df106["片区编码"].replace("宁波", "0201", inplace=True)
        df106["片区编码"].replace("舟山北仑", "0202", inplace=True)
        df106["片区编码"].replace("北三县", "0203", inplace=True)
        df106["片区编码"].replace("南三县", "0204", inplace=True)
        df106["片区编码"].replace("杭州省级", "0301", inplace=True)
        df106["片区编码"].replace("杭州市级", "0302", inplace=True)  #
        df106["片区编码"].replace("嘉湖", "0303", inplace=True)  #
        df106["片区编码"].replace("南京吴珏", "0401", inplace=True)  # 0129修改
        df106["片区编码"].replace("南京高跃", "0402", inplace=True)
        df106["片区编码"].replace("南通李国旺", "0501", inplace=True)
        df106["片区编码"].replace("盐城", "0503", inplace=True)
        df106["片区编码"].replace("连云港", "0504", inplace=True)
        df106["片区编码"].replace("上海1", "0601", inplace=True)
        df106["片区编码"].replace("上海2", "0602", inplace=True)
        df106["片区编码"].replace("上海3", "0603", inplace=True)  # 0129修改
        df106["片区编码"].replace("苏州", "0701", inplace=True)
        df106["片区编码"].replace("苏州市郊", "0702", inplace=True)
        df106["片区编码"].replace("扬州", "0801", inplace=True)
        df106["片区编码"].replace("泰州", "0802", inplace=True)
        df106["片区编码"].replace("徐州", "0901", inplace=True)  #
        df106["片区编码"].replace("宿迁", "0902", inplace=True)
        df106["片区编码"].replace("淮安", "0903", inplace=True)
        df106["片区编码"].replace("常州", "1001", inplace=True)
        df106["片区编码"].replace("镇江", "1002", inplace=True)
        df106["片区编码"].replace("绍兴龚群波", "1101", inplace=True)
        df106["片区编码"].replace("金衢", "1102", inplace=True)  #
        df106["片区编码"].replace("无锡1", "1201", inplace=True)
        df106["片区编码"].replace("无锡2", "1202", inplace=True)
        df106["片区编码"].replace("业务拓展部", "1301", inplace=True)
        df106["片区编码"].replace("调拨", "1901", inplace=True)
        df106["片区编码"].replace("配件", "1902", inplace=True)
        df106["片区编码"].replace("生化分销", "1903", inplace=True)

        df107 = df106.drop(["片区编码_x", "片区"], axis=1)  # 删列

        # df106.to_excel(excel_writer="D:/78.xlsx",
        #      sheet_name="处理",
        #    index=False);
        # df191.to_excel(excel_writer="D:/79.xlsx",
        #                sheet_name="处理",
        #                index=False);

        ####俩表新建部门字段，修改为编码，然后合并0707

        ####拼接接团账龄
        # df107 = pd.merge(df19,df106, right_index=True, left_index=True);

        df107 = pd.merge(df191, df107, how='left', on=['片区编码', '部门', '公司','客户']);  # 完全相同合并
        df107["三个月以上应收账款"] = df107["Unnamed: 10"]

        df1071 = df107.drop(["Unnamed: 10"], axis=1)  # 删列 df108 改1071

        dfa108 = df1071[df1071["片区"] == "调拨检验所"]
        dfa108["部门"].replace("调拨", "调拨-检验所", inplace=True)

        dfa110a = df1071[df1071["片区"] != "调拨检验所"]
        df108 = pd.concat([dfa110a, dfa108], ignore_index=True)

        # df108.to_excel(excel_writer="D:/二稿测试0902.xlsx",
        #            sheet_name="处理",
        #          index=False);
        # dfa110a.to_excel(excel_writer="D:/合并测试/二稿测试11295.xlsx",
        #                sheet_name="处理",
        #                index=False);
        #######以上是试剂和仪器收款拼接完毕,下面开始增加部门小计（调拨类的分开统计）
        dfa200 = df108[df108["部门"] != "调拨-检验所"]
        df201 = dfa200[dfa200["部门"] != "调拨"]  # 去除部门为调拨，配件，生化分销
        df201 = df201[df201["部门"] != "配件"]  # 去除部门为调拨，配件，生化分销
        df202 = df201[df201["部门"] != "生化分销"]  #

        # 0203去除部门中的检验所后合计
        df203 = df202[df202["片区"] != "一部检验所"]
        df204 = df203[df203["片区"] != "二部检验所"]
        df205 = df204[df204["片区"] != "三部检验所"]
        df206 = df205[df205["片区"] != "四部检验所"]
        df207 = df206[df206["片区"] != "五部检验所"]
        df208 = df207[df207["片区"] != "六部检验所"]
        df209 = df208[df208["片区"] != "七部检验所"]
        df210 = df209[df209["片区"] != "八部检验所"]
        df211 = df210[df210["片区"] != "九部检验所"]
        df212 = df211[df211["片区"] != "十部检验所"]
        df213 = df212[df212["片区"] != "十一部检验所"]
        df214 = df213[df213["片区"] != "十二部检验所"]

        df20 = df214.groupby(["部门"], as_index=False)[
            "本月试剂销售", "本月试剂收款", "期初应收余额", "期末应收余额", "本期仪器收款", "本期仪器销售", "三个月以上应收账款"].sum();
        df20["片区"] = df20["部门"]
        df20["片区编码"] = df20["部门"]
        df20["公司"] = df20["部门"]

        df20["片区"].replace("一部", "诊断一部", inplace=True)
        df20["片区"].replace("二部", "诊断二部", inplace=True)
        df20["片区"].replace("三部", "诊断三部", inplace=True)
        df20["片区"].replace("四1部", "诊断四1部", inplace=True)
        df20["片区"].replace("四2部", "诊断四2部", inplace=True)
        df20["片区"].replace("五1部", "诊断五1部", inplace=True)
        df20["片区"].replace("五2部", "诊断五2部", inplace=True)
        df20["片区"].replace("六部", "诊断六部", inplace=True)
        df20["片区"].replace("七部", "诊断七部", inplace=True)
        df20["片区"].replace("八部", "诊断八部", inplace=True)
        df20["片区"].replace("九部", "诊断九部", inplace=True)
        df20["片区"].replace("十部", "诊断十部", inplace=True)
        df20["片区"].replace("十一部", "诊断十一部", inplace=True)
        df20["片区"].replace("十二部", "诊断十二部", inplace=True)
        df20["片区"].replace("业务拓展部", "业务拓展部", inplace=True)  # 0129修改

        df20["片区编码"].replace("一部", "0199", inplace=True)
        df20["片区编码"].replace("二部", "0299", inplace=True)
        df20["片区编码"].replace("三部", "0399", inplace=True)
        df20["片区编码"].replace("四1部", "04019", inplace=True)
        df20["片区编码"].replace("四2部", "04029", inplace=True)
        df20["片区编码"].replace("五1部", "05019", inplace=True)
        df20["片区编码"].replace("五2部", "05049", inplace=True)
        df20["片区编码"].replace("六部", "0699", inplace=True)
        df20["片区编码"].replace("七部", "0799", inplace=True)
        df20["片区编码"].replace("八部", "0899", inplace=True)
        df20["片区编码"].replace("九部", "0999", inplace=True)
        df20["片区编码"].replace("十部", "1099", inplace=True)
        df20["片区编码"].replace("十一部", "1199", inplace=True)
        df20["片区编码"].replace("十二部", "1299", inplace=True)
        df20["片区编码"].replace("业务拓展部", "1399", inplace=True)

        df20["公司"].replace("一部", "总计", inplace=True)
        df20["公司"].replace("二部", "总计", inplace=True)
        df20["公司"].replace("三部", "总计", inplace=True)
        df20["公司"].replace("四1部", "总计", inplace=True)
        df20["公司"].replace("四2部", "总计", inplace=True)
        df20["公司"].replace("五1部", "总计", inplace=True)
        df20["公司"].replace("五2部", "总计", inplace=True)
        df20["公司"].replace("六部", "总计", inplace=True)
        df20["公司"].replace("七部", "总计", inplace=True)
        df20["公司"].replace("八部", "总计", inplace=True)
        df20["公司"].replace("九部", "总计", inplace=True)
        df20["公司"].replace("十部", "总计", inplace=True)
        df20["公司"].replace("十一部", "总计", inplace=True)
        df20["公司"].replace("十二部", "总计", inplace=True)
        df20["公司"].replace("业务拓展部", "总计", inplace=True)
        df21 = pd.concat([df108, df20], ignore_index=True)
        df211 = df21[df21["部门"] != ""]

        df22 = df211.sort_values(by=['片区编码'], axis=0, ascending=True)  # 行排序

        ####初步组合完毕

        ####剥离调拨部门

        df23 = df22[df22["部门"] != "一部"]
        df231 = df23[df23["部门"] != "二部"]
        df232 = df231[df231["部门"] != "三部"]
        df233 = df232[df232["部门"] != "四1部"]
        df234 = df233[df233["部门"] != "五1部"]
        df235 = df234[df234["部门"] != "六部"]
        df236 = df235[df235["部门"] != "七部"]
        df237 = df236[df236["部门"] != "八部"]
        df238 = df237[df237["部门"] != "九部"]
        df239 = df238[df238["部门"] != "十部"]
        df240 = df239[df239["部门"] != "十一部"]
        df241 = df240[df240["部门"] != "十二部"]
        df242 = df241[df241["部门"] != "业务拓展部"]
        df243 = df242[df242["部门"] != "四2部"]
        df244 = df243[df243["部门"] != "五2部"]
        df2411 = df244[df244["片区"] != "调拨检验所"]
        df2411["片区编码"] = df2411["部门"]

        # df241.to_excel(excel_writer="D:/合并测试/二稿测试1.xlsx",sheet_name="处理",index=False);

        df2411["片区编码"].replace("调拨", "1900", inplace=True)
        df2411["片区编码"].replace("配件", "1900", inplace=True)
        df2411["片区编码"].replace("生化分销", "1900", inplace=True)

        df24 = df2411.groupby(["片区编码", "公司"], as_index=False)[
            "本月试剂销售", "本月试剂收款", "期初应收余额", "期末应收余额", "本期仪器收款", "本期仪器销售", "三个月以上应收账款"].sum();

        df24["部门"] = df24["片区编码"]
        df24["片区"] = df24["片区编码"]
        # df24["公司"] = df24["片区编码"]

        df24["部门"].replace("1900", "调拨", inplace=True)
        df24["片区"].replace("1900", "调拨部", inplace=True)
        # df24["公司"].replace("1900", "总计", inplace=True)

        # df24.to_excel(excel_writer="D:/2020新表/调拨合计测试.xlsx",
        #             sheet_name="处理",
        #           index=False);

        df25 = df22[df22["部门"] != "调拨"]
        df26 = df25[df25["部门"] != "配件"]
        df27 = df26[df26["部门"] != "生化分销"]

        df28 = pd.concat([df27, df24], ignore_index=True)

        # 0203新增调拨合计

        df2811 = df24.groupby(["片区编码"], as_index=False)[
            "本月试剂销售", "本月试剂收款", "期初应收余额", "期末应收余额", "本期仪器收款", "本期仪器销售", "三个月以上应收账款"].sum();
        df2811["公司"] = df2811["片区编码"]
        df2811["片区"] = df2811["片区编码"]
        df2811["公司"].replace("1900", "总计", inplace=True)
        df2811["片区编码"].replace("1900", "1999", inplace=True)
        df2811["片区"].replace("1900", "调拨部", inplace=True)

        df281 = pd.concat([df28, df2811], ignore_index=True)

        #####调拨整合完毕,引用模板为df28修改三部,九部,十一部,汇总地区

        df63 = df281.sort_values(by=['片区编码'], axis=0, ascending=True)  # 行排序
        print(df63)

        #####数据组合完毕,开始表格列排序

        df63["部门名称"] = df63["部门"]
        df63["地区"] = df63["片区"]
        df63["负责人"] = df63["片区"]
        df63["地区编码"] = df63["片区编码"]
        df63["本期试剂销售"] = df63["本月试剂销售"]
        df63["本期试剂收款"] = df63["本月试剂收款"]
        df63["期初余额"] = df63["期初应收余额"]
        df63["期末余额"] = df63["期末应收余额"]
        df63["本月仪器收款"] = df63["本期仪器收款"]
        df63["本月仪器销售"] = df63["本期仪器销售"]
        df63["超三个月应收账款"] = df63["三个月以上应收账款"]

        df64 = df63.drop(["部门", "片区", "片区编码", "本月试剂销售", "本月试剂收款", "期初应收余额", "期末应收余额", "本期仪器收款", "本期仪器销售", "三个月以上应收账款"],
                         axis=1)  # 删列

        # df64.to_excel(excel_writer="D:/09021.xlsx",  #####710
        #              sheet_name="处理",
        #              );
        ####开始修改片区负责人
        df64["负责人"].replace("温州1", "葛瑞", inplace=True)
        df64["负责人"].replace("温州2", "潘磊", inplace=True)
        df64["负责人"].replace("台州1", "唐惠", inplace=True)
        df64["负责人"].replace("台州2", "胡文魁", inplace=True)
        df64["负责人"].replace("丽水", "方汝泼", inplace=True)
        df64["负责人"].replace("诊断一部", "郭德春", inplace=True)

        df64["负责人"].replace("宁波", "丁玲", inplace=True)
        df64["负责人"].replace("舟山北仑", "高大勇", inplace=True)
        df64["负责人"].replace("北三县", "陆金耀", inplace=True)
        df64["负责人"].replace("南三县", "吴燕江", inplace=True)
        df64["负责人"].replace("诊断二部", "余顶峰", inplace=True)

        df64["负责人"].replace("杭州省级", "姜立民", inplace=True)
        df64["负责人"].replace("杭州市级", "沈剑芳", inplace=True)  #
        df64["负责人"].replace("嘉湖", "阮芳", inplace=True)  #
        df64["负责人"].replace("诊断三部", "毛存亮", inplace=True)

        df64["负责人"].replace("南京高跃", "高跃", inplace=True)
        df64["负责人"].replace("南京吴珏", "吴珏", inplace=True)
        df64["负责人"].replace("诊断四2部", "高跃", inplace=True)
        df64["负责人"].replace("诊断四1部", "吴珏", inplace=True)

        df64["负责人"].replace("南通李国旺", "李国旺", inplace=True)
        df64["负责人"].replace("盐城", "岑潭泽 潘前进", inplace=True)
        df64["负责人"].replace("连云港", "胡士艳 姜健", inplace=True)
        df64["负责人"].replace("诊断五1部", "李国旺", inplace=True)
        df64["负责人"].replace("诊断五2部", "胡士艳", inplace=True)

        df64["负责人"].replace("上海1", "邬幼波", inplace=True)
        df64["负责人"].replace("上海2", "汤俊", inplace=True)
        df64["负责人"].replace("上海3", "黄帅", inplace=True)
        df64["负责人"].replace("诊断六部", "毛存亮", inplace=True)

        df64["负责人"].replace("苏州", "陈凯", inplace=True)
        df64["负责人"].replace("苏州市郊", "吕楠", inplace=True)
        df64["地区"].replace("苏州", "苏州市区", inplace=True)
        df64["地区"].replace("苏州市郊", "苏州郊县", inplace=True)
        df64["负责人"].replace("诊断七部", "全英娜", inplace=True)

        df64["负责人"].replace("扬州", "邹海洵", inplace=True)
        df64["负责人"].replace("泰州", "金英明", inplace=True)
        df64["负责人"].replace("诊断八部", "金英明 邹海洵", inplace=True)

        df64["负责人"].replace("徐州", "唐维洲 于博", inplace=True)  #

        df64["负责人"].replace("宿迁", "赵晨阳 王涛", inplace=True)
        df64["负责人"].replace("淮安", "赵晨阳 白虹", inplace=True)
        df64["负责人"].replace("诊断九部", "吴蓓", inplace=True)

        df64["负责人"].replace("常州", "梅晓虹", inplace=True)
        df64["负责人"].replace("镇江", "于亚惠", inplace=True)
        df64["负责人"].replace("诊断十部", "梅晓虹", inplace=True)

        df64["地区"].replace("绍兴龚群波", "绍兴", inplace=True)
        df64["负责人"].replace("绍兴龚群波", "龚群波", inplace=True)

        df64["负责人"].replace("金衢", "胡迪锋", inplace=True)  #
        df64["负责人"].replace("诊断十一部", "李征", inplace=True)

        df64["负责人"].replace("无锡1", "裘涌", inplace=True)
        df64["负责人"].replace("无锡2", "张立伟 赵飞", inplace=True)
        df64["负责人"].replace("诊断十二部", "邬幼波", inplace=True)

        df64["负责人"].replace("调拨部", "孙婷婷", inplace=True)
        df64["负责人"].replace("业务拓展部", "叶仲华", inplace=True)

        df64["地区编码"].replace("一部检验所", "5001", inplace=True)
        df64["地区编码"].replace("二部检验所", "5002", inplace=True)
        df64["地区编码"].replace("三部检验所", "5003", inplace=True)
        df64["地区编码"].replace("四部检验所", "5004", inplace=True)
        df64["地区编码"].replace("五部检验所", "5005", inplace=True)
        df64["地区编码"].replace("六部检验所", "5006", inplace=True)
        df64["地区编码"].replace("七部检验所", "5007", inplace=True)
        df64["地区编码"].replace("八部检验所", "5008", inplace=True)
        df64["地区编码"].replace("九部检验所", "5009", inplace=True)
        df64["地区编码"].replace("十部检验所", "5010", inplace=True)
        df64["地区编码"].replace("十一部检验所", "5011", inplace=True)
        df64["地区编码"].replace("十二部检验所", "5012", inplace=True)
        df64["地区编码"].replace("调拨检验所", "5013", inplace=True)

        df65 = df64.sort_values(by=['地区编码'], axis=0, ascending=True)  # 行排序

        df66 = df65[df65["部门名称"] != "诊断二部"]
        df67 = df66[df66["部门名称"] != "诊断一部"]

        # 0203去掉一部检验所等显示
        df68 = df67[df67["地区"] != "一部检验所"]
        df69 = df68[df68["地区"] != "二部检验所"]
        df70 = df69[df69["地区"] != "三部检验所"]
        df71 = df70[df70["地区"] != "四部检验所"]
        df72 = df71[df71["地区"] != "五部检验所"]
        df73 = df72[df72["地区"] != "六部检验所"]
        df74 = df73[df73["地区"] != "七部检验所"]
        df75 = df74[df74["地区"] != "八部检验所"]
        df76 = df75[df75["地区"] != "九部检验所"]
        df77 = df76[df76["地区"] != "十部检验所"]
        df78 = df77[df77["地区"] != "十一部检验所"]
        df79 = df78[df78["地区"] != "十二部检验所"]
        df80 = df79[df79["地区"] != "调拨检验所"]



        # 全表合计 df82
        df81 = df80[df80["公司"] == "总计"]
        df82 = df81.groupby(["公司"], as_index=False)[
            "本期试剂销售", "本期试剂收款", "期初余额", "期末余额", "本月仪器收款", "本月仪器销售", "超三个月应收账款"].sum();

        df82["公司"].replace("总计", "合计", inplace=True)

        df82["地区编码"] = df82["公司"]
        df82["地区"] = df82["公司"]
        df82["地区编码"].replace("总计", "9999", inplace=True)
        df82["地区"].replace("总计", "合计", inplace=True)

        df83 = pd.concat([df80, df82], ignore_index=True)

        # 全表排序
        df84 = df83.groupby(["公司", "地区", "地区编码","客户"], as_index=False)[
            "本期试剂销售", "本期试剂收款", "期初余额", "期末余额", "本月仪器收款", "本月仪器销售", "超三个月应收账款"].sum();

        df85 = df84.sort_values(by=['地区编码'], axis=0, ascending=True)  # 行排序

        ####20200902新需求 加入仪器应收################################################

        df0101 = df100['客户'].str.split('-| ', expand=True);
        df0102 = pd.merge(df0101, df100, right_index=True, left_index=True);

        df0102["部门"] = df0102[0]
        df0102["片区"] = df0102[1]
        df0102["片区编码_x"] = df0102[2]
        df01021 = df0102.rename(columns={'客户': '编码'});
        df01021["客户"] = df01021[7]


        df0103 = df01021[df01021["项目"] == "003 仪器"]
        df0104 = df0103[df0103["公司"] != "海尔施集团"]

        df0105 = df0104.groupby(["客户","公司"], as_index=False)["余额"].sum();

        # df86 = pd.concat([df85, df0105], ignore_index=True)
        df0106 = df0105.rename(columns={'余额': '仪器余额'})  # 把原来的 客户 命名为 客户名称
        df86 = pd.merge(df85,df0106,how='outer',on=['客户','公司']);





        df87 = df86.sort_values(by=['地区编码'], axis=0, ascending=True)  # 行排序


        #########################################考核表完成~0711

        # df64.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630 df64是第二稿模板

        # sheets.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630

        df87.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                         filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                                    (
                                                                                        "Microsoft Excel 97-20003 文件",
                                                                                        "*.xls")],
                                                                         defaultextension=".xlsx"));

        tkinter.messagebox.showinfo("运行结果", "考核表2020明细导出成功！");

    except Exception as error:
        tm.showerror(title="煎饼提示前方路堵",
                     message="请检查提交源文件是否正确 '" + str(error) + "'.",
                     detail=traceback.format_exc())



def appendStr41():#客户分列单体账龄
 try:
    tkinter.messagebox.showinfo("提醒", "请选择没有项目的金蝶单体公司账龄表");
    df16 = pd.read_excel(tkinter.filedialog.askopenfilename());
    df18 = df16['客户'].str.split('-| ',expand=True);
    print(df18)
    df19 = pd.merge(df18, df16, right_index=True, left_index=True);
    df20 = df19.drop(df19.columns[[8,9]], axis=1)
    df22 = df20.dropna();

    #df22.to_excel(excel_writer="d:/分列前端数据361.xlsx",
     #                      sheet_name="测试1",
      #                    index=False);


    df22['客户'] = df22[7]
    df22['部门'] = df22[0]
    df22['片区'] = df22[1]
    df22['客户编码'] = df22[3]

    print(df22)
    df23 = df22.groupby(["部门","片区","客户","客户编码"])["过期", "Unnamed: 6", "Unnamed: 7", "Unnamed: 8", "Unnamed: 9", "Unnamed: 10"].sum();
    c_df = pd.DataFrame(df23)
    c_df.reset_index(inplace=True)  # 取消合并

   # df23.to_excel(excel_writer="d:/分列前端数据2.xlsx",
   #               sheet_name="测试1",
    #              index=False);

    df24 = df23.rename(columns={'过期': '1-30','Unnamed: 6':'31-60','Unnamed: 7':'61-90','Unnamed: 8':'91-120',
                                 'Unnamed: 9':'121-150','Unnamed: 10':'151-'});

    df24["部门"].replace("一部", "01部", inplace=True)
    df24["部门"].replace("二部", "02部", inplace=True)
    df24["部门"].replace("三部", "03部", inplace=True)
    df24["部门"].replace("四1部", "04部1", inplace=True)
    df24["部门"].replace("四2部", "04部2", inplace=True)
    df24["部门"].replace("五1部", "05部1", inplace=True)
    df24["部门"].replace("五2部", "05部2", inplace=True)
    df24["部门"].replace("六部", "06部", inplace=True)
    df24["部门"].replace("七部", "07部", inplace=True)
    df24["部门"].replace("八部", "08部", inplace=True)
    df24["部门"].replace("九部", "09部", inplace=True)
    df24["部门"].replace("十部", "10部", inplace=True)
    df24["部门"].replace("十一部", "11部", inplace=True)
    df24["部门"].replace("十二部", "12部", inplace=True)
    df24["片区"].replace("", "无", inplace=True)


    df24 = df24[df24["片区"] != "无"]
    df24["余额"] = df24["1-30"] + df24["31-60"] + df24["61-90"] + df24["91-120"] + df24["121-150"] + df24["151-"]





    df25 = df24.sort_values(by=['部门','片区'], axis=0, ascending=True)
    print(df25)

    df26 = df25.groupby(["部门", "片区", "客户编码","客户","余额"])[ "1-30", "31-60", "61-90", "91-120", "121-150", "151-"].sum();
    c_df = pd.DataFrame(df26)
    c_df.reset_index(inplace=True)  # 取消合并



   # df26.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630
    df26.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"));
    tkinter.messagebox.showinfo("运行结果","分列并导出成功！");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())



def appendStr43():  # K3科目表做转换 dhy
     try:
         tkinter.messagebox.showinfo("提醒", "请选择要转换的表格");
         df1 = pd.read_excel(tkinter.filedialog.askopenfilename());

         df2 = df1.loc[df1['科目名称'].str.contains('海尔施')]
         df3 = df1.loc[df1['科目名称'].str.contains('恒奇诊断')]
         df4 = df1.loc[df1['科目名称'].str.contains('宁波大红鹰药业')]
         df5 = df1.loc[df1['科目名称'].str.contains('宁波高新区大红鹰医药进出口有限公司')]

         df6 = df1.loc[df1['科目名称'].str.contains('海壹生物科技')]
         df7 = df1.loc[df1['科目名称'].str.contains('江苏恒奇医药有限公司')]
         df8 = df1.loc[df1['科目名称'].str.contains('沭阳恒昌')]
         df9 = df1.loc[df1['科目名称'].str.contains('金华强盛生物科技')]
         df10 = df1.loc[df1['科目名称'].str.contains('宁波美晶')]

         df40 = pd.concat([df2, df3, df4, df5, df6, df7, df8, df9, df10],
                          ignore_index=True)  # 组合
         df41 = df40.sort_values(by=['科目代码'], axis=0, ascending=True)  # 行排序

         df42 = df41[df41["科目代码"] != "1511.01"]
         df43 = df42[df42["科目代码"] != "1511.02"]
         df44 = df43[df43["科目代码"] != "1511.03"]
         df45 = df44[df44["科目代码"] != "4001.02"]
         df46 = df45[df45["科目名称"] != "[01.02.16.001]宁波大红鹰药业股份有限公司医务室"]

         d11 = df46[df46["科目名称"] != "预收账款1"]
         d21 = d11[d11["科目名称"] != "预付往来款"]
         d31 = d21[d21["科目名称"] != "预付往来款1"]
         d41 = d31[d31["科目名称"] != "预付工程款"]
         d51 = d41[d41["科目名称"] != "预付工程款2"]
         d61 = d51[d51["科目名称"] != "预付设备款"]
         d71 = d61[d61["科目名称"] != "预付设备款3"]
         d81 = d71[d71["科目名称"] != "待摊费用"]




         df50 = d81.groupby(["科目代码"], as_index=False)[
             "期末借方余额", "期末贷方余额"].sum();

         df50["类型"] = df50["科目代码"]
         df50["类型"].replace("1122", "dhy内部", inplace=True)
         df50["类型"].replace("1221.01", "dhy内部", inplace=True)
         df50["类型"].replace("2202.01", "dhy内部", inplace=True)
         df50["类型"].replace("2241.01", "dhy内部", inplace=True)

         df50["科目代码"].replace("1221.01", "1221", inplace=True)
         df50["科目代码"].replace("2202.01", "2202", inplace=True)
         df50["科目代码"].replace("2241.01", "2241", inplace=True)

         #######以上是内部交易表

         ######以下是全部交易

         df100 = df1[df1['科目代码'].isin(["1122", "2202.01", "2202.02", "2202.03", "2203", "1221.01", "1221.02", "1221.03",
                                       "2241.01", "2241.02", "2241.03", "2241.04", "2241.05", "2241.06", "2241.07",
                                       "2241.08",
                                       "2241.09.03", "2241.10", "2241.99.01", "2241.99.02", "2241.99.04", "2241.99.05",
                                       "2203", "1123.01.01", "1123.02.02", "1123.03.03", "1123.04.02", "1123.04.03",
                                       "1123.04.04", "1123.04.06", "1123.04.10", "1123.04.11"])]

         df101 = df100[df100["科目名称"] != "应收账款"]
         df10101 = df101[df101["科目名称"] != "客户单位往来"]
         df10102 = df10101[df10101["科目名称"] != "押金"]
         df10103 = df10102[df10102["科目名称"] != "个人往来"]
         df10104 = df10103[df10103["科目名称"] != "供应商往来"]
         df10105 = df10104[df10104["科目名称"] != "预收账款"]
         df10106 = df10105[df10105["科目名称"] != "单位往来"]
         df10107 = df10106[df10106["科目名称"] != "应付工程、设备款"]
         df10107a = df10107[df10107["科目名称"] != "预收账款"]

         d1 = df10107a[df10107a["科目名称"] != "预收账款1"]
         d2 = d1[d1["科目名称"] != "预付往来款"]
         d3 = d2[d2["科目名称"] != "预付往来款1"]
         d4 = d3[d3["科目名称"] != "预付工程款"]
         d5 = d4[d4["科目名称"] != "预付工程款2"]
         d6 = d5[d5["科目名称"] != "预付设备款"]
         d7 = d6[d6["科目名称"] != "预付设备款3"]
         d8 = d7[d7["科目名称"] != "待摊费用"]





         # df110 = pd.concat([df100, df101, df102, df103, df104, df105],
         #                 ignore_index=True)  # 组合
         df112 = d8.sort_values(by=['科目代码'], axis=0, ascending=True)  # 行排序

         df112["科目代码"].replace("1123.01.01", "1123", inplace=True)
         df112["科目代码"].replace("1123.02.02", "1123", inplace=True)
         df112["科目代码"].replace("1123.03.03", "1123", inplace=True)
         df112["科目代码"].replace("1123.04.02", "1123", inplace=True)
         df112["科目代码"].replace("1123.04.03", "1123", inplace=True)
         df112["科目代码"].replace("1123.04.04", "1123", inplace=True)
         df112["科目代码"].replace("1123.04.06", "1123", inplace=True)
         df112["科目代码"].replace("1123.04.10", "1123", inplace=True)
         df112["科目代码"].replace("1123.04.11", "1123", inplace=True)

         df112["科目代码"].replace("1221.01", "1221", inplace=True)
         df112["科目代码"].replace("1221.02", "1221", inplace=True)
         df112["科目代码"].replace("1221.03", "1221", inplace=True)

         df112["科目代码"].replace("2202.01", "2202", inplace=True)
         df112["科目代码"].replace("2202.02", "2202", inplace=True)
         df112["科目代码"].replace("2202.03", "2202", inplace=True)

         df112["科目代码"].replace("2241.01", "2241", inplace=True)
         df112["科目代码"].replace("2241.03", "2241", inplace=True)
         df112["科目代码"].replace("2241.04", "2241", inplace=True)
         df112["科目代码"].replace("2241.05", "2241", inplace=True)
         df112["科目代码"].replace("2241.06", "2241", inplace=True)
         df112["科目代码"].replace("2241.07", "2241", inplace=True)
         df112["科目代码"].replace("2241.09.03", "2241", inplace=True)
         df112["科目代码"].replace("2241.10", "2241", inplace=True)
         df112["科目代码"].replace("2241.99.01", "2241", inplace=True)
         df112["科目代码"].replace("2241.99.02", "2241", inplace=True)
         df112["科目代码"].replace("2241.99.04", "2241", inplace=True)
         df112["科目代码"].replace("2241.99.05", "2241", inplace=True)

         df113 = df112.groupby(["科目代码"], as_index=False)[
             "期末借方余额", "期末贷方余额"].sum();

         df113["类型"] = df113["科目代码"]
         df113["类型"].replace("1122", "全部", inplace=True)
         df113["类型"].replace("1123", "全部", inplace=True)
         df113["类型"].replace("1221", "全部", inplace=True)
         df113["类型"].replace("2202", "全部", inplace=True)
         df113["类型"].replace("2203", "全部", inplace=True)
         df113["类型"].replace("2241", "全部", inplace=True)

         ######以上是合并交易

         df199 = pd.merge(df113, df50, how='left', on=['科目代码']);  # 完全相同合并
         # df200 = pd.concat([df50, df113],ignore_index=True)  # 组合

         df200 = df199.fillna(0);
         # df6["Unnamed: 14"] = df6["Unnamed: 14"].astype("float64");  # 改变格式
         df200["外部期末借方余额"] = df200["期末借方余额_x"] - df200["期末借方余额_y"]
         df200["外部期末贷方余额"] = df200["期末贷方余额_x"] - df200["期末贷方余额_y"]

         # df26.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630
         df200.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                          filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                                     ("Microsoft Excel 97-20003 文件",
                                                                                      "*.xls")],
                                                                          defaultextension=".xlsx"), sheet_name="表1");
         tkinter.messagebox.showinfo("运行结果", "导出成功！");
     except Exception as error:
         tm.showerror(title="煎饼提示前方路堵",
                      message="请检查提交源文件是否正确 '" + str(error) + "'.",
                      detail=traceback.format_exc())



def appendStr44():  # K3科目表做转换 qs
    try:
        tkinter.messagebox.showinfo("提醒", "请选择要转换的表格");
        df1 = pd.read_excel(tkinter.filedialog.askopenfilename());

        df2 = df1.loc[df1['科目名称'].str.contains('海尔施')]
        df3 = df1.loc[df1['科目名称'].str.contains('恒奇诊断')]
        df4 = df1.loc[df1['科目名称'].str.contains('宁波大红鹰药业')]
        df5 = df1.loc[df1['科目名称'].str.contains('宁波高新区大红鹰医药进出口有限公司')]

        df6 = df1.loc[df1['科目名称'].str.contains('海壹生物科技')]
        df7 = df1.loc[df1['科目名称'].str.contains('江苏恒奇医药有限公司')]
        df8 = df1.loc[df1['科目名称'].str.contains('沭阳恒昌')]
        df9 = df1.loc[df1['科目名称'].str.contains('金华强盛生物科技')]
        df10 = df1.loc[df1['科目名称'].str.contains('宁波美晶')]

        df40 = pd.concat([df2, df3, df4, df5, df6, df7, df8, df9, df10],
                         ignore_index=True)  # 组合
        df41 = df40.sort_values(by=['科目代码'], axis=0, ascending=True)  # 行排序

        df42 = df41[df41["科目代码"] != "6001.01"]

        df42["科目代码"].replace("2202.02", "2202", inplace=True)
        df42["科目代码"].replace("2202.01", "2202", inplace=True)
        df42["科目代码"].replace("2241.21", "2241", inplace=True)



        df50 = df42.groupby(["科目代码"], as_index=False)[
            "期末借方余额", "期末贷方余额"].sum();

        df50["类型"] = df50["科目代码"]
        df50["类型"].replace("1122", "qs内部", inplace=True)
        df50["类型"].replace("1221.01", "qs内部", inplace=True)
        df50["类型"].replace("2202.01", "qs内部", inplace=True)
        df50["类型"].replace("2241.21", "qs内部", inplace=True)
        df50["类型"].replace("2202.02", "qs内部", inplace=True)
        df50["类型"].replace("2202", "qs内部", inplace=True)
        df50["类型"].replace("2241", "qs内部", inplace=True)

        #######以上是内部交易表
        df100 = df1[df1['科目代码'].isin(["1122", "1123.02","1221.01","1221.02","2202.01",
                                      "2202.02","2241.16","2241.21","2241.23","2241.24",
                                      "2241.25","2241.27"])]

        df101 = df100[df100["科目名称"] != "应收账款"]
        df102 = df101[df101["科目名称"] != "应付账款"]
        df103 = df102[df102["科目名称"] != "应付账款—暂估"]
        df104 = df103[df103["科目名称"] != "内部往来"]
        df105 = df104[df104["科目名称"] != "内部往来"]
        df106 = df105[df105["科目名称"] != "保证金"]
        df107 = df106[df106["科目名称"] != "外部往来"]
        df108 = df107[df107["科目名称"] != "代扣社保"]
        df109 = df108[df108["科目名称"] != "员工往来款"]
        df110 = df109[df109["科目名称"] != "待摊费用"]
        df111 = df110[df110["科目名称"] != "预收账款"]




        df112 = df111.sort_values(by=['科目代码'], axis=0, ascending=True)  # 行排序
        df112["科目代码"].replace("1123.02", "1123", inplace=True)
        df112["科目代码"].replace("1221.01", "1221", inplace=True)
        df112["科目代码"].replace("1221.02", "1221", inplace=True)
        df112["科目代码"].replace("2202.01", "2202", inplace=True)
        df112["科目代码"].replace("2202.02", "2202", inplace=True)
        df112["科目代码"].replace("2241.16", "2241", inplace=True)
        df112["科目代码"].replace("2241.21", "2241", inplace=True)
        df112["科目代码"].replace("2241.23", "2241", inplace=True)
        df112["科目代码"].replace("2241.24", "2241", inplace=True)
        df112["科目代码"].replace("2241.25", "2241", inplace=True)
        df112["科目代码"].replace("2241.27", "2241", inplace=True)

        df113 = df112.groupby(["科目代码"], as_index=False)[
            "期末借方余额", "期末贷方余额"].sum();

        df113["类型"] = df113["科目代码"]
        df113["类型"].replace("1122", "全部", inplace=True)
        df113["类型"].replace("1123", "全部", inplace=True)
        df113["类型"].replace("1221", "全部", inplace=True)
        df113["类型"].replace("2202", "全部", inplace=True)
        df113["类型"].replace("2203", "全部", inplace=True)
        df113["类型"].replace("2241", "全部", inplace=True)

        ######以上是合并交易

        df199 = pd.merge(df113, df50, how='left', on=['科目代码']);  # 完全相同合并
        # df200 = pd.concat([df50, df113],ignore_index=True)  # 组合
        df200=df199.fillna(0)
        # df6["Unnamed: 14"] = df6["Unnamed: 14"].astype("float64");  # 改变格式
        df200["外部期末借方余额"] = df200["期末借方余额_x"].astype("float64") - df200["期末借方余额_y"].astype("float64")
        df200["外部期末贷方余额"] = df200["期末贷方余额_x"].astype("float64") - df200["期末贷方余额_y"].astype("float64")







        # df26.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630
        df200.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                         filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                                    ("Microsoft Excel 97-20003 文件",
                                                                                     "*.xls")],
                                                                         defaultextension=".xlsx"), sheet_name="表1");
        tkinter.messagebox.showinfo("运行结果", "导出成功！");
    except Exception as error:
        tm.showerror(title="煎饼提示前方路堵",
                     message="请检查提交源文件是否正确 '" + str(error) + "'.",
                     detail=traceback.format_exc())


def appendStr45():  # K3科目表做转换jck
    try:
        tkinter.messagebox.showinfo("提醒", "请选择要转换的表格");
        df1 = pd.read_excel(tkinter.filedialog.askopenfilename());

        df2 = df1.loc[df1['科目名称'].str.contains('海尔施')]
        df3 = df1.loc[df1['科目名称'].str.contains('恒奇诊断')]
        df4 = df1.loc[df1['科目名称'].str.contains('宁波大红鹰药业')]
        df5 = df1.loc[df1['科目名称'].str.contains('宁波高新区大红鹰医药进出口有限公司')]

        df6 = df1.loc[df1['科目名称'].str.contains('海壹生物科技')]
        df7 = df1.loc[df1['科目名称'].str.contains('江苏恒奇医药有限公司')]
        df8 = df1.loc[df1['科目名称'].str.contains('沭阳恒昌')]
        df9 = df1.loc[df1['科目名称'].str.contains('金华强盛生物科技')]
        df10 = df1.loc[df1['科目名称'].str.contains('宁波美晶')]

        df40 = pd.concat([df2, df3, df4, df5, df6, df7, df8, df9, df10],
                         ignore_index=True)  # 组合
        df41 = df40.sort_values(by=['科目代码'], axis=0, ascending=True)  # 行排序

        df42 = df41[df41["科目代码"] != "4001.02"]

        df42["科目代码"].replace("2202.01", "2202", inplace=True)
        df42["科目代码"].replace("2241.01", "2241", inplace=True)




        df50 = df42.groupby(["科目代码"], as_index=False)[
            "期末借方余额", "期末贷方余额"].sum();

        df50["类型"] = df50["科目代码"]
        df50["类型"].replace("2241", "jck内部", inplace=True)

        #######以上是内部交易表
        df100 = df1[df1['科目代码'].isin(["1122", "1221.01", "2202.01", "2203",
                                      "2241.01", "2241.04", "2241.05", "2241.06",
                                      "2241.07", "2241.99.05"])]

        df101 = df100[df100["科目名称"] != "应收账款"]
        df102 = df101[df101["科目名称"] != "供应商往来"]
        df103 = df102[df102["科目名称"] != "单位往来"]
        df104 = df103[df103["科目名称"] != "预收账款"]

        df112 = df104.sort_values(by=['科目代码'], axis=0, ascending=True)  # 行排序


        df112["科目代码"].replace("1221.01", "1221", inplace=True)
        df112["科目代码"].replace("2202.01", "2202", inplace=True)

        df112["科目代码"].replace("2241.01", "2241", inplace=True)
        df112["科目代码"].replace("2241.04", "2241", inplace=True)
        df112["科目代码"].replace("2241.05", "2241", inplace=True)
        df112["科目代码"].replace("2241.06", "2241", inplace=True)
        df112["科目代码"].replace("2241.07", "2241", inplace=True)
        df112["科目代码"].replace("2241.99.05", "2241", inplace=True)

        df113 = df112.groupby(["科目代码"], as_index=False)[
            "期末借方余额", "期末贷方余额"].sum();

        df113["类型"] = df113["科目代码"]
        df113["类型"].replace("1122", "全部", inplace=True)
        df113["类型"].replace("1123", "全部", inplace=True)
        df113["类型"].replace("1221", "全部", inplace=True)
        df113["类型"].replace("2202", "全部", inplace=True)
        df113["类型"].replace("2203", "全部", inplace=True)
        df113["类型"].replace("2241", "全部", inplace=True)

        ######以上是合并交易

        df199 = pd.merge(df113, df50, how='left', on=['科目代码']);  # 完全相同合并
        # df200 = pd.concat([df50, df113],ignore_index=True)  # 组合
        df200 = df199.fillna(0)
        # df6["Unnamed: 14"] = df6["Unnamed: 14"].astype("float64");  # 改变格式
        df200["外部期末借方余额"] = df200["期末借方余额_x"].astype("float64") - df200["期末借方余额_y"].astype("float64")
        df200["外部期末贷方余额"] = df200["期末贷方余额_x"].astype("float64") - df200["期末贷方余额_y"].astype("float64")




        # df26.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630
        df200.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                         filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                                    ("Microsoft Excel 97-20003 文件",
                                                                                     "*.xls")],
                                                                         defaultextension=".xlsx"), sheet_name="表1");
        tkinter.messagebox.showinfo("运行结果", "导出成功！");
    except Exception as error:
        tm.showerror(title="煎饼提示前方路堵",
                     message="请检查提交源文件是否正确 '" + str(error) + "'.",
                     detail=traceback.format_exc())

#################################################################################################################################3
########发票认证
def appendStr101():
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
    #df8.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
     #                                                              filetypes=[("Microsoft Excel文件", "*.xlsx"),
      #                                                                        ("Microsoft Excel 97-2003 文件", "*.xls")],
      #                                                             defaultextension=".xls"),index=False);

    tkinter.messagebox.showinfo("运行结果","需认证整理成功!");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())

def appendStr102():
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
    #df8.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
     #                                                              filetypes=[("Microsoft Excel文件", "*.xlsx"),
      #                                                                        ("Microsoft Excel 97-2003 文件", "*.xls")],
       #                                                            defaultextension=".xls"),index=False);

    tkinter.messagebox.showinfo("运行结果","需认证整理成功!");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())

def appendStr103():
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
    # df81=df8.drop(df8.columns[[0,1]],axis=1)
    #########

    df300 = pd.merge(df2, df, how='left', on=['发票号码']);
    df301=df300[df300["是否勾选(是/否)_y"] != "否"]

    #df302=df301.drop(["是否勾选(是/否)_x", "是否勾选(是/否)_y", "发票代码", "开票日期","税额","有效抵扣税额","销方名称","销方税号","金额","用途","Unnamed: 1","Unnamed: 2","Unnamed: 3","Unnamed: 4","Unnamed: 5","Unnamed: 6","Unnamed: 7"], axis=1)  # 删列

    print(df8)

    df8.to_excel("本次认证整理文件" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                    index=False)  # 自动输出
    df301.to_excel("本次认证失败明细" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx", sheet_name="sheet1",
                 index=False)  # 自动输出


    tkinter.messagebox.showinfo("运行结果","需认证整理成功!");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())



def appendStr81():
 try:
    # tkinter.messagebox.showinfo("提醒", "请选择开票明细表");
    # df = pd.read_excel("d:/data.xlsx", sheet_name="sheet1");  #
    df = pd.read_excel(tkinter.filedialog.askopenfilename());
    df1 = df.drop(df.columns[[[[[[[[[[[0, 3, 4, 5, 7, 8, 10, 11, 12, 13, 17]]]]]]]]]]], axis=1);
    df2 = df1.drop(df1.index[[[[0, 1, 2, 3]]]], axis=0);
    df3 = df2[df2["Unnamed: 9"] != "小计"];
    df4 = df3[df3["Unnamed: 9"] != "商品名称"];
    df5 = df4.dropna(how="all");
    df6 = df5.fillna(method='pad');
    df6["Unnamed: 14"] = df6["Unnamed: 14"].astype("float64");  # 改变格式
    df6["Unnamed: 16"] = df6["Unnamed: 16"].astype("float64");

    df10 = df6

    #df10 = pd.read_excel("d:/data4.xlsx", sheet_name="测试1")
    df10["Unnamed: 14"] = df10["Unnamed: 14"].astype("float64");
    df10["Unnamed: 16"] = df10["Unnamed: 16"].astype("float64");
    df100 = df10.rename(columns={'Unnamed: 1': '发票号码'});
    df11 = df100.groupby(["发票号码","Unnamed: 2","Unnamed: 6"])["Unnamed: 14","Unnamed: 16"].sum() ;
    # df12 = df11['发票合计'] = df11.apply(lambda x: x.sum(), axis=1);
    # df13 = df12.rename(columns={'Unnamed: 1':'发票号码'});

    print(df11)

    # df20 = df100.groupby("发票号码")["Unnamed: 14", "Unnamed: 16"].sum();
    #
    # df21 = pd.merge(df11, df20,how='left');  # 完全相同合并
    df11["合计"]=df11["Unnamed: 14"]+df11["Unnamed: 16"]
    #tkinter.filedialog.asksaveasfile(mode='w',
     #   defaultextension='.txt',

    #df12.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb',defaultextension='*.xlsx',));#指定位置另存为630

    df11.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"));

    tkinter.messagebox.showinfo("运行结果","客户发票税额汇总成功!");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())


def appendStr82():#大邱转换
 try:
    tkinter.messagebox.showinfo("提醒", "请选择要转换的表格");
    df1 = pd.read_excel(tkinter.filedialog.askopenfilename());
    c_df = pd.DataFrame(df1)
    c_df.reset_index(inplace=True)  # 取消合并

    df2 = c_df[c_df["index"] == " "]
    df3 = df2.drop(["index"],axis = 1)#删列


    df14 = c_df[c_df["index"] == 0]
    df4 = c_df[c_df["index"] == 171]
    df5 = c_df[c_df["index"] == 172]
    df6 = c_df[c_df["index"] == 173]
    df7 = c_df[c_df["index"] == 174]
    df8 = c_df[c_df["index"] == 175]
    df9 = c_df[c_df["index"] == 176]
    df10 = c_df[c_df["index"] == 177]
    df11 = c_df[c_df["index"] == 178]
    df12 = c_df[c_df["index"] == 179]
    df13 = c_df[c_df["index"] == 180]

    df20 = pd.concat([df14,df4, df5, df6, df7,df8,df9,df10,df11,df12,df13])
    df21 = df20.drop(["index"], axis=1)  # 删列

    df22 = pd.concat([df3, df21])

    df23 = df22.rename(
        columns={'Unnamed: 4': '01海尔施生物', 'Unnamed: 6': '02海壹生物', 'Unnamed: 8': '03浙江医疗', 'Unnamed: 10': '04基因',
                 'Unnamed: 12': '23中翰金诺', 'Unnamed: 14': '06医药有限', 'Unnamed: 16': '07上海诊断', 'Unnamed: 18': '08海壹',
                 'Unnamed: 20': '09医学检验所'
            , 'Unnamed: 22': '10供应链', 'Unnamed: 24': '12大红鹰', 'Unnamed: 26': '13大红鹰进出口', 'Unnamed: 28': '14恒奇',
                 'Unnamed: 30': '17恒奇诊断'
            , 'Unnamed: 32': '18上海器械', 'Unnamed: 34': '19金华强盛', 'Unnamed: 36': '20美晶小合并', 'Unnamed: 38': '合计',
                 'Unnamed: 1': '科目', 'Unnamed: 2': '科目'});

    #df2 = c_df.drop([[[[[[[[[[3,4,5,6,7,8,9,10,11,12]]]]]]]]]],axis=0)  # 删行
    df24 = df23.T  # 行列互换

    df25 = df24.drop(df24.index[0], axis=0)





   # df26.to_excel(excel_writer=tkinter.filedialog.asksaveasfile(mode='wb', defaultextension='.xlsx'));  # 指定位置另存为630
    df25.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               ("Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"),sheet_name="表1");
    tkinter.messagebox.showinfo("运行结果","导出成功！");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
                  message="请检查提交源文件是否正确 '" + str(error) + "'.",
                  detail=traceback.format_exc())

def appendStr83():  # 月末销售-英克-金蝶核对
 try:
      tkinter.messagebox.showinfo("提醒", "请选择金蝶科目项目余额表源文件");
      # 加入应收账款借方本年发生额

      df150 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      c_df = pd.DataFrame(df150)
      c_df.reset_index(inplace=True)

      df151 = df150.drop([0], axis=0)

      df152 = df151.drop(["index","科目代码","科目名称","期初余额","Unnamed: 5","Unnamed: 7","本年累计","Unnamed: 9","期末余额","Unnamed: 11"], axis=1)  # 删列
      df153 = df152.rename(columns={'项目代码': '金蝶编码'});

      tkinter.messagebox.showinfo("提醒", "请选择英克销售明细源文件");
      # 加入英克销售明细表

      df70 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径
      df71 = df70.rename(columns={'客户ID': '英克代码'});


      #tkinter.messagebox.showinfo("提醒", "请选择金蝶映射表源文件");
      # 加入金蝶映射表

      #df160 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      df160 = pd.read_excel('http://nbhealth.eicp.net:14692/khysb.xlsx')

      df161= df160.loc[df160['英克代码'].str.contains('y')]  # 7-15模糊查询,但单元格不能为0为空  （ #print(df4[df4['客户'].isin(['宁波海尔施医学检验所有限公司'])])）
      df161["金诺类型"]='否'







      df162 = pd.merge(df160, df161, how='left', on=['英克代码']);  # 完全相同合并，忽略没有的货品ID(没有how)
      df163 =df162[df162["金诺类型"] != "否"]

      df164 = df163.drop(["   _x", "集团名称_x", "集团编码_x", "类型_x", "   _y", "客户名称_y", "集团名称_y", "金蝶编码_y", "集团编码_y", "类型_y","金诺类型","客户名称_x"],axis=1)  # 删列
      df165 = df164.rename(columns={'金蝶编码_x': '金蝶编码'});
      # df160['英克代码'].astype("float64")
      # df162 = pd.merge(df71, df161, how='outer', on=['英克代码']);  # 完全相同合并，忽略没有的货品ID(没有how)



      ###组合两表

      df165['英克代码']=df165['英克代码'].astype("float64") #改列格式字符改数值

      # 过滤英克编码重复
      df166 = df165.drop_duplicates(['英克代码'])

      df167 = pd.merge(df71, df166, how='inner', on=['英克代码']);  # 完全相同合并，忽略没有的货品ID(没有how)


      # 和金蝶销售明细组合
      # 先汇总df167     # 拼接金蝶销售

      df168 = df167.groupby(['金蝶编码'])["金额"].sum();



      df170 = pd.merge(df153, df168, how='inner', on=['金蝶编码'])

      df171 = df170.rename(columns={'金额': '英克金额','本期发生额':'金蝶金额'});
      df171['差异']=df171['金蝶金额']-df171['英克金额']





      #加入销售发票
      tkinter.messagebox.showinfo("提醒", "请选择销售开票源文件");
      df = pd.read_excel(tkinter.filedialog.askopenfilename());
      df1 = df.drop(df.columns[[[[[[[[[[[0, 3, 4, 5, 7, 8, 10, 11, 12, 13, 17]]]]]]]]]]], axis=1);
      df2 = df1.drop(df1.index[[[[0, 1, 2, 3]]]], axis=0);
      df3 = df2[df2["Unnamed: 9"] != "小计"];
      df4 = df3[df3["Unnamed: 9"] != "商品名称"];
      df5 = df4.dropna(how="all");
      df6 = df5.fillna(method='pad');
      df6["Unnamed: 14"] = df6["Unnamed: 14"].astype("float64");  # 改变格式
      df6["Unnamed: 16"] = df6["Unnamed: 16"].astype("float64");

      df10 = df6


      df10["Unnamed: 14"] = df10["Unnamed: 14"].astype("float64");
      df10["Unnamed: 16"] = df10["Unnamed: 16"].astype("float64");
      df11 = df10.rename(columns={'Unnamed: 2': '客户'});
      df12 = df11.groupby("客户")["Unnamed: 14", "Unnamed: 16"].sum();
      df12['发票合计'] = df12.apply(lambda x: x.sum(), axis=1);


      # df13 = df12.rename(columns={'Unnamed: 2': '客户'});
      print(df12)

      #组合2表df171
      df172 = pd.merge(df171, df12, how='outer', on=['客户'])
      #客户名称排序
      df173 = df172.sort_values(by=['客户'], axis=0, ascending=True)  # 行排序

      df174 = df173.fillna(0)
      df174['金蝶金额'].astype("float64")
      df174['发票合计'].astype("float64")


      df174['差异2']=df174['金蝶金额'].astype("float64")-df174['发票合计'].astype("float64")
      df174['差异2'].astype("float64")
      df175 = df174.drop(["Unnamed: 14", "Unnamed: 16"], axis=1)  # 删

      df175.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"), sheet_name="表1");
      tkinter.messagebox.showinfo("运行结果", "导出成功！");

 except Exception as error:
      tm.showerror(title="煎饼提示前方路堵",
              message="请检查提交源文件是否正确 '" + str(error) + "'.",
               detail=traceback.format_exc())




def appendStr84():  # 信用控制带英克ID
 try:
      tkinter.messagebox.showinfo("提醒", "请选择控制表源文件");


      df1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径
      df2 = pd.read_excel('http://nbhealth.eicp.net:14692/khysb.xlsx')
      df3 = pd.merge(df1, df2, how='left', on=['客户名称'])

      df4 =  df3.drop_duplicates(['英克代码'])
      df5 = df4.drop(df4.columns[[[[[[[[[[[[[[[[[[[[[[[[0, 1, 2, 3, 4, 5, 6,7, 8, 9,10, 11, 12, 13,14,15,16,17,18,20,21,22]]]]]]]]]]]]]]]]]]]]]]]], axis=1);
      df5['英克ID']=df5['英克代码']+'，'
      df6 = df5.drop(df5.columns[0],axis=1);

      print('开始写入txt文件...')
      df6.to_csv('本月客户信用英克ID.txt', header=None, sep=',', index=False)  # 写入，逗号分隔
      print('文件写入成功!')

      # df5.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
      #                                                               filetypes=[("Microsoft Excel文件", "*.xlsx"),
      #                                                                          ("Microsoft Excel 97-20003 文件", "*.xls")],
      #                                                               defaultextension=".xlsx"), sheet_name="表1",index=False);
      tkinter.messagebox.showinfo("运行结果", "导出成功！");

 except Exception as error:
      tm.showerror(title="煎饼提示前方路堵",
              message="请检查提交源文件是否正确 '" + str(error) + "'.",
               detail=traceback.format_exc())


def appendStr85():  # 月末毛利分析
 try:
      tkinter.messagebox.showinfo("提醒", "请选择生物源文件");
      df1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      df2 = df1.groupby(['客户','货品ID','三级分类'])["销售成本","基本单位数量","价额"].sum();
      df2.reset_index(inplace=True)  # 取消合并


      d2 = df1.groupby(['客户', '货品ID', '三级分类','保管账名称'])["销售成本", "基本单位数量", "价额"].sum();
      d2.reset_index(inplace=True)  # 取消合并
      #####单体
      da2 = df1.groupby(['客户', '货品ID', '三级分类'])["销售成本", "基本单位数量", "价额"].sum();
      da2.reset_index(inplace=True)  # 取消合并


      tkinter.messagebox.showinfo("提醒", "请选择浙江源文件");
      df3 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      df4 = df3.groupby(['客户', '货品ID', '三级分类'])["销售成本", "基本单位数量", "价额"].sum();
      df4.reset_index(inplace=True)  # 取消合并

      d4 = df3.groupby(['客户', '货品ID', '三级分类','保管账名称'])["销售成本", "基本单位数量", "价额"].sum();
      d4.reset_index(inplace=True)  # 取消合并
      #####单体
      da4 = df3.groupby(['客户', '货品ID', '三级分类'])["销售成本", "基本单位数量", "价额"].sum();
      da4.reset_index(inplace=True)  # 取消合并

      tkinter.messagebox.showinfo("提醒", "请选择上海诊断源文件");
      df5 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      df6 = df5.groupby(['客户', '货品ID', '三级分类'])["销售成本", "基本单位数量", "价额"].sum();
      df6.reset_index(inplace=True)  # 取消合并

      d6 = df5.groupby(['客户', '货品ID', '三级分类','保管账名称'])["销售成本", "基本单位数量", "价额"].sum();
      d6.reset_index(inplace=True)  # 取消合并
      #####单体
      da6 = df5.groupby(['客户', '货品ID', '三级分类'])["销售成本", "基本单位数量", "价额"].sum();
      da6.reset_index(inplace=True)  # 取消合并

      tkinter.messagebox.showinfo("提醒", "请选择上海器械源文件");
      df7 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      df8 = df7.groupby(['客户', '货品ID', '三级分类'])["销售成本", "基本单位数量", "价额"].sum();
      df8.reset_index(inplace=True)  # 取消合并

      d8 = df7.groupby(['客户', '货品ID', '三级分类','保管账名称'])["销售成本", "基本单位数量", "价额"].sum();
      d8.reset_index(inplace=True)  # 取消合并
      #####单体
      da8 = df7.groupby(['客户', '货品ID', '三级分类'])["销售成本", "基本单位数量", "价额"].sum();
      da8.reset_index(inplace=True)  # 取消合并

      tkinter.messagebox.showinfo("提醒", "请选择恒奇诊断源文件");
      df9 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      df10 = df9.groupby(['客户', '货品ID', '三级分类'])["销售成本", "基本单位数量", "价额"].sum();
      df10.reset_index(inplace=True)  # 取消合并

      d10 = df9.groupby(['客户', '货品ID', '三级分类','保管账名称'])["销售成本", "基本单位数量", "价额"].sum();
      d10.reset_index(inplace=True)  # 取消合并
      #####单体
      da10 = df9.groupby(['客户', '货品ID', '三级分类'])["销售成本", "基本单位数量", "价额"].sum();
      da10.reset_index(inplace=True)  # 取消合并
   ######组合1-5



      ##################

      df20 = pd.concat([df2,df4,df6,df8,df10],ignore_index=True)
      d20 = pd.concat([d2, d4, d6, d8, d10], ignore_index=True)    #d20加了保管账     ####11-10导出单体明细
   ######关联方区分  生物加恒奇诊断
      ####10-19 去掉伯乐指控,010501质控试剂（代理）

      df40 = d20.loc[d20['保管账名称'].str.contains('海尔施生物')]
      d40 = df40[df40["三级分类"] != '010501质控试剂（代理）']
      d41 = d20.loc[d20['保管账名称'].str.contains('01配件保管账')]

      #01配件保管账

      df41 = d20.loc[d20['保管账名称'].str.contains('江苏恒奇保管账')]
      df42 = df41[df41["三级分类"] == '010501质控试剂（代理）']
      df43 =pd.concat([d40,d41,df42],ignore_index=True)

      # d42 = df42[df42["客户"] != '海尔施生物医药股份有限公司']
      # d43 = d42[d42["客户"] != '江苏恒奇诊断产品有限公司']
      # d44 = d43[d43["客户"] != '江苏恒奇诊断产品有限公司（固定资产）']

      df44 = df43.groupby(['货品ID'])["销售成本","基本单位数量"].sum();   #去掉客户
      df44.reset_index(inplace=True)  # 取消合并
      df44['采购进价'] = (df44['销售成本'].astype("float64")) / (df44['基本单位数量'].astype("float64"))

      # df44.to_excel(excel_writer="d:/1020.xlsx",
      #            sheet_name="测试1",
      #            index = False);

      # df44 = df43.rename(columns={'销售成本': '关联销售成本','基本单位数量':'关联基本单位数量','价额':'关联价格'});
      df45 = df44.drop(["基本单位数量","销售成本"], axis=1)
    ########关联方标注"是"

      df49 = df20[df20["客户"]!= '宁波海尔施基因科技有限公司']
      df50 = df49.loc[df49['客户'].str.contains('海尔施')]
      df50["关联方"] = '是'
      df51 = df49.loc[df49['客户'].str.contains('恒奇诊断')]
      df51["关联方"] = '是'
      df52 = df49.loc[df49['客户'].str.contains('宁波美晶')]
      df52["关联方"] = '是'
      df53 = df49.loc[df49['客户'].str.contains('海壹生物')]
      df53["关联方"] = '是'
      df54 = df49.loc[df49['客户'].str.contains('强盛生物')]
      df54["关联方"] = '是'
      df55 = df49.loc[df49['客户'].str.contains('杭州金诺')]
      df55["关联方"] = '是'


      df60 = pd.concat([df50,df51,df52,df53,df54,df55],ignore_index=True)

      df61 = pd.merge(df20, df60, how='outer', on=['客户', '货品ID', '三级分类','销售成本', '基本单位数量', '价额']);  # 完全相同合并，忽略没有的货品ID(没有how)
      #
      # ####
      #
      df62 = pd.merge(df61, df45, how='outer', on=['货品ID']);

      # df62["采购成本"]=df62["基本单位数量"].astype("float64")*df62["采购进价"].astype("float64")
      # df62["毛利"] = (df62["价额"].astype("float64")-df62["基本单位数量"].astype("float64")*df62["采购进价"].astype("float64"))
      # df62["毛利率"]=(df62["价额"].astype("float64")-df62["基本单位数量"].astype("float64")*df62["采购进价"].astype("float64"))/df62["价额"].astype("float64")


      #####按照项目汇总
      tkinter.messagebox.showinfo("提醒", "请选择分类规则源文件");
      df63 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      df64 = pd.merge(df62, df63, how='left', on=['三级分类']);

      # df69 = df64.groupby(['大类'])["销售成本", "基本单位数量", "价额", "毛利"].sum();
      # df70 = df64.groupby(['二级分类'])["销售成本", "基本单位数量", "价额","毛利"].sum();
      # # df70.reset_index(inplace=True)  # 取消合并
      #
      #
      # df71=pd.concat([df69,df70],ignore_index=True)

      df65 = df64.fillna(0)
      ####################
      # df64["采购进价"].fillna(df64["销售成本"]/df64["基本单位数量"])
      ##############
      # df65 = df64[df64["采购进价"] != '']
      # # d66 = df65[df65["采购进价"] != '']
      #
      # df66 = df64[df64["采购进价"] == ' ']
      #
      # df66["采购进价"]=df66["销售成本"]/df66["基本单位数量"]
      #
      # df67 = pd.concat([df65, df66], ignore_index=True)

      df66 = df65[df65["采购进价"] == 0]
      df67 = df65[df65["采购进价"] != 0]
      #修改匹配时为int -int



      df66["采购进价"]=df66["销售成本"]/df66["基本单位数量"]

      df68 = pd.concat([df67, df66], ignore_index=True)

      df68["采购成本"]=df68["采购进价"]*df68["基本单位数量"]
      #######开始汇总非关联方

      #######单体公司11-10分析



      # 3-医药 新加d12


      # df69 = df68[df68["关联方"] != "是"]
      #
      # df70 = df69.groupby(['大类','二级分类'])["价额","基本单位数量","采购进价"].sum();   #去掉客户
      #
      #
      #
      # df70.reset_index(inplace=True)  # 取消合并
      # df70["采购成本"]=df70["基本单位数量"]*df70["采购进价"]




      df68.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"), sheet_name="成本分析明细");
      tkinter.messagebox.showinfo("运行结果", "明细导出成功！");

      ###d2包含保管账开始
      # 1-生物d2

      sw49 = da2[da2["客户"] != '宁波海尔施基因科技有限公司']
      sw50 = sw49.loc[sw49['客户'].str.contains('海尔施')]
      sw50["关联方"] = '是'
      sw51 = sw49.loc[sw49['客户'].str.contains('恒奇诊断')]
      sw51["关联方"] = '是'
      sw52 = sw49.loc[sw49['客户'].str.contains('宁波美晶')]
      sw52["关联方"] = '是'
      sw53 = sw49.loc[sw49['客户'].str.contains('海壹生物')]
      sw53["关联方"] = '是'
      sw54 = sw49.loc[sw49['客户'].str.contains('强盛生物')]
      sw54["关联方"] = '是'
      sw55 = sw49.loc[sw49['客户'].str.contains('杭州金诺')]
      sw55["关联方"] = '是'

      sw60 = pd.concat([sw50, sw51, sw52, sw53, sw54, sw55], ignore_index=True)
      # sw601 = d2.drop(["基本单位数量"], axis=1)
      sw61 = pd.merge(da2, sw60, how='outer', on=['客户', '货品ID', '三级分类', '销售成本','价额','基本单位数量']);
      sw61["毛利"] = sw61["价额"] - sw61["销售成本"]
      sw62 = sw61.fillna({"关联方": "否"})
      sw63 = sw62.groupby(['三级分类'])["销售成本","价额","毛利"].sum();
      sw64 = pd.merge(sw62, df63, how='left', on=['三级分类']);#63是输入
      sw65 = sw64.groupby(['大类'])["销售成本","价额","毛利"].sum();

      write = pd.ExcelWriter("生物毛利分析" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx")

      sw62.to_excel(write, sheet_name='单体明细')
      sw63.to_excel(write, sheet_name='三级分类')
      sw65.to_excel(write, sheet_name='大类')
      write.save()
      write.close()
      time.sleep(5)
      # tkinter.messagebox.showinfo("运行结果", "完成！");

      # 2-浙江d4
      zj49 = da4[da4["客户"] != '宁波海尔施基因科技有限公司']
      zj50 = zj49.loc[zj49['客户'].str.contains('海尔施')]
      zj50["关联方"] = '是'
      zj51 = zj49.loc[zj49['客户'].str.contains('恒奇诊断')]
      zj51["关联方"] = '是'
      zj52 = zj49.loc[zj49['客户'].str.contains('宁波美晶')]
      zj52["关联方"] = '是'
      zj53 = zj49.loc[zj49['客户'].str.contains('海壹生物')]
      zj53["关联方"] = '是'
      zj54 = zj49.loc[zj49['客户'].str.contains('强盛生物')]
      zj54["关联方"] = '是'
      zj55 = zj49.loc[zj49['客户'].str.contains('杭州金诺')]
      zj55["关联方"] = '是'

      zj60 = pd.concat([zj50, zj51, zj52, zj53, zj54, zj55], ignore_index=True)
      # zj601 = d2.drop(["基本单位数量"], axis=1)
      zj61 = pd.merge(da4, zj60, how='outer', on=['客户', '货品ID', '三级分类', '销售成本', '价额','基本单位数量']);
      zj61["毛利"] = zj61["价额"] - zj61["销售成本"]
      zj62 = zj61.fillna({"关联方": "否"})
      zj63 = zj62.groupby(['三级分类'])["销售成本", "价额", "毛利"].sum();
      zj64 = pd.merge(zj62, df63, how='left', on=['三级分类']);  # 63是输入
      zj65 = zj64.groupby(['大类'])["销售成本", "价额", "毛利"].sum();

      write = pd.ExcelWriter("浙江毛利分析" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx")

      zj62.to_excel(write, sheet_name='单体明细')
      zj63.to_excel(write, sheet_name='三级分类')
      zj65.to_excel(write, sheet_name='大类')
      write.save()
      write.close()
      time.sleep(5)
      # 4-上海器械d8
      shqx49 = da8[da8["客户"] != '宁波海尔施基因科技有限公司']
      shqx50 = shqx49.loc[shqx49['客户'].str.contains('海尔施')]
      shqx50["关联方"] = '是'
      shqx51 = shqx49.loc[shqx49['客户'].str.contains('恒奇诊断')]
      shqx51["关联方"] = '是'
      shqx52 = shqx49.loc[shqx49['客户'].str.contains('宁波美晶')]
      shqx52["关联方"] = '是'
      shqx53 = shqx49.loc[shqx49['客户'].str.contains('海壹生物')]
      shqx53["关联方"] = '是'
      shqx54 = shqx49.loc[shqx49['客户'].str.contains('强盛生物')]
      shqx54["关联方"] = '是'
      shqx55 = shqx49.loc[shqx49['客户'].str.contains('杭州金诺')]
      shqx55["关联方"] = '是'

      shqx60 = pd.concat([shqx50, shqx51, shqx52, shqx53, shqx54, shqx55], ignore_index=True)
      # shqx601 = d2.drop(["基本单位数量"], axis=1)
      shqx61 = pd.merge(da8, shqx60, how='outer', on=['客户', '货品ID', '三级分类', '销售成本','价额','基本单位数量']);
      shqx61["毛利"] = shqx61["价额"] - shqx61["销售成本"]
      shqx62 = shqx61.fillna({"关联方": "否"})
      shqx63 = shqx62.groupby(['三级分类'])["销售成本", "价额", "毛利"].sum();
      shqx64 = pd.merge(shqx62, df63, how='left', on=['三级分类']);  # 63是输入
      shqx65 = shqx64.groupby(['大类'])["销售成本", "价额", "毛利"].sum();

      write = pd.ExcelWriter("上海器械毛利分析" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx")

      shqx62.to_excel(write, sheet_name='单体明细')
      shqx63.to_excel(write, sheet_name='三级分类')
      shqx65.to_excel(write, sheet_name='大类')
      write.save()
      write.close()
      time.sleep(5)
      # 5-上海诊断d6
      shzd49 = da6[da6["客户"] != '宁波海尔施基因科技有限公司']
      shzd50 = shzd49.loc[shzd49['客户'].str.contains('海尔施')]
      shzd50["关联方"] = '是'
      shzd51 = shzd49.loc[shzd49['客户'].str.contains('恒奇诊断')]
      shzd51["关联方"] = '是'
      shzd52 = shzd49.loc[shzd49['客户'].str.contains('宁波美晶')]
      shzd52["关联方"] = '是'
      shzd53 = shzd49.loc[shzd49['客户'].str.contains('海壹生物')]
      shzd53["关联方"] = '是'
      shzd54 = shzd49.loc[shzd49['客户'].str.contains('强盛生物')]
      shzd54["关联方"] = '是'
      shzd55 = shzd49.loc[shzd49['客户'].str.contains('杭州金诺')]
      shzd55["关联方"] = '是'

      shzd60 = pd.concat([shzd50, shzd51, shzd52, shzd53, shzd54, shzd55], ignore_index=True)
      # shzd601 = d2.drop(["基本单位数量"], axis=1)
      shzd61 = pd.merge(da6, shzd60, how='outer', on=['客户', '货品ID', '三级分类', '销售成本','价额','基本单位数量']);
      shzd61["毛利"] = shzd61["价额"] - shzd61["销售成本"]
      shzd62 = shzd61.fillna({"关联方": "否"})
      shzd63 = shzd62.groupby(['三级分类'])["销售成本", "价额", "毛利"].sum();
      shzd64 = pd.merge(shzd62, df63, how='left', on=['三级分类']);  # 63是输入
      shzd65 = shzd64.groupby(['大类'])["销售成本", "价额", "毛利"].sum();

      write = pd.ExcelWriter("上海诊断毛利分析" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx")

      shzd62.to_excel(write, sheet_name='单体明细')
      shzd63.to_excel(write, sheet_name='三级分类')
      shzd65.to_excel(write, sheet_name='大类')
      write.save()
      write.close()
      time.sleep(5)
      # 6-恒奇诊断d10
      jshq49 = da10[da10["客户"] != '宁波海尔施基因科技有限公司']
      jshq50 = jshq49.loc[jshq49['客户'].str.contains('海尔施')]
      jshq50["关联方"] = '是'
      jshq51 = jshq49.loc[jshq49['客户'].str.contains('恒奇诊断')]
      jshq51["关联方"] = '是'
      jshq52 = jshq49.loc[jshq49['客户'].str.contains('宁波美晶')]
      jshq52["关联方"] = '是'
      jshq53 = jshq49.loc[jshq49['客户'].str.contains('海壹生物')]
      jshq53["关联方"] = '是'
      jshq54 = jshq49.loc[jshq49['客户'].str.contains('强盛生物')]
      jshq54["关联方"] = '是'
      jshq55 = jshq49.loc[jshq49['客户'].str.contains('杭州金诺')]
      jshq55["关联方"] = '是'

      jshq60 = pd.concat([jshq50, jshq51, jshq52, jshq53, jshq54, jshq55], ignore_index=True)
      # jshq601 = d2.drop(["基本单位数量"], axis=1)
      jshq61 = pd.merge(da10, jshq60, how='outer', on=['客户', '货品ID', '三级分类', '销售成本','价额','基本单位数量']);
      jshq61["毛利"] = jshq61["价额"] - jshq61["销售成本"]
      jshq62 = jshq61.fillna({"关联方": "否"})
      jshq63 = jshq62.groupby(['三级分类'])["销售成本", "价额", "毛利"].sum();
      jshq64 = pd.merge(jshq62, df63, how='left', on=['三级分类']);  # 63是输入
      jshq65 = jshq64.groupby(['大类'])["销售成本", "价额", "毛利"].sum();

      write = pd.ExcelWriter("恒奇诊断毛利分析" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx")

      jshq62.to_excel(write, sheet_name='单体明细')
      jshq63.to_excel(write, sheet_name='三级分类')
      jshq65.to_excel(write, sheet_name='大类')
      write.save()
      write.close()
      tkinter.messagebox.showinfo("运行结果", "单体毛利分析都已生成在小程序左右！");
 except Exception as error:
      tm.showerror(title="煎饼提示前方路堵",
              message="请检查提交源文件是否正确 '" + str(error) + "'.",
               detail=traceback.format_exc())

def appendStr86():  #暂估核对组
 try:
      tkinter.messagebox.showinfo("提醒", "请选择金蝶暂估项目核算表源文件");

      df150 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      c_df = pd.DataFrame(df150)
      c_df.reset_index(inplace=True)

      df151 = df150.drop([0], axis=0)

      df152 = df151.drop(
         ["index", "科目代码", "科目名称", "期初余额", "Unnamed: 5", "Unnamed: 7", "本年累计", "Unnamed: 9","本期发生额"],
         axis=1)  # 删列  , "Unnamed: 11"是贷方 ,期末余额 是借方
      df153 = df152.rename(columns={'项目代码': '供应商编码','Unnamed: 11':'金蝶暂估贷方','期末余额':'金蝶暂估借方'});
      # df153['供应商编码'] = df153['供应商编码'].astype("float64")  # 改列格式字符改数值     1201

      # 加入英克发票未到
      tkinter.messagebox.showinfo("提醒", "请选择英克发票未到源文件");
      df70 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径
      df71 = df70.rename(columns={'供应商ID': '英克编码'});

      df72 = df71.groupby(["英克编码"], as_index=False)["未结算成本金额"].sum();

      # tkinter.messagebox.showinfo("提醒", "请选择金蝶映射表源文件");
      # 加入金蝶映射表

      # df160 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      df160 = pd.read_excel('http://nbhealth.eicp.net:14692/gysysb.xlsx')

      df161 = df160.loc[
         df160['英克编码'].str.contains('y')]  # 7-15模糊查询,但单元格不能为0为空  （ #print(df4[df4['客户'].isin(['宁波海尔施医学检验所有限公司'])])）
      df161["金诺类型"] = '否'

      df162 = pd.merge(df160, df161, how='left', on=['英克编码']);  # 完全相同合并，忽略没有的ID(没有how)
      df163 = df162[df162["金诺类型"] != "否"]

      df164 = df163.drop(
         ["   _x","   _y", "供应商名称_y", "供应商编码_y","金诺类型",
          "供应商名称_x"], axis=1)  # 删列
      df165 = df164.rename(columns={'供应商编码_x': '供应商编码'});
      df165['英克编码'] = df165['英克编码'].astype("float64")  # 改列格式字符改数值

      ##先拼接英克
      df166 = pd.merge(df72, df165, how='inner', on=['英克编码'])
      # df166['供应商编码'] = df166['供应商编码'].astype("float64")  # 改列格式字符改数值 1201
      # df167 = df166.drop_duplicates(['供应商编码'])
      df167 = df166.groupby("供应商编码")["未结算成本金额"].sum();


      ###两表合并

      df168 = pd.merge(df153, df167, how='outer', on=['供应商编码'])

      df169 =df168[df168["供应商"] != "合计"]

      df170 = df169.fillna(0)
      df170['暂估差异'] = df170['金蝶暂估贷方'].astype("float64") - df170['金蝶暂估借方'].astype("float64") - df170['未结算成本金额'].astype("float64")
      df170['暂估差异'].astype("float64")
      df170.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                               "Microsoft Excel 97-20003 文件", "*.xls")],
                                                                    defaultextension=".xlsx"), sheet_name="表1");
      tkinter.messagebox.showinfo("运行结果", "暂估核对导出成功！");

 except Exception as error:
      tm.showerror(title="煎饼提示前方路堵",
              message="请检查提交源文件是否正确 '" + str(error) + "'.",
               detail=traceback.format_exc())



def appendStr101():  #2020英克毛利
 try:
      tkinter.messagebox.showinfo("提醒", "请放入生物采购明细单");
      df1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      # sw65 = sw64[~sw64['三级分类'].str.contains('仪器')]  # 取反
      #
      # df2 = df1[~df1['供应商'].str.contains('海尔施')]  # 取反
      df2 = df1[df1['供应商'].str.contains('海尔施')==False]  # 取反
      df3 = df2[~df2['供应商'].str.contains('恒奇诊断')]  # 取反

      df4 = df1[df1["供应商"] == "宁波海尔施基因科技有限公司"]
      df5 = pd.concat([df3,df4], ignore_index=True)

      # df4 = df3[~df3['供应商'].str.contains('宁波美晶')]  # 取反
      # df5 = df4[~df4['供应商'].str.contains('海壹生物')]  # 取反
      # df6 = df5[~df5['供应商'].str.contains('强盛生物')]  # 取反
      # df7 = df6[~df6['供应商'].str.contains('杭州金诺')]  # 取反

      # df8 = df7.loc[df7.reset_index().groupby(['货品ID'])['成本单价'].idxmax()]      #取成本单价最高得一笔
      # df9 = df8.dropna(how='all')  #全部缺失就删除

      tkinter.messagebox.showinfo("提醒", "请放入上海采购明细单");
      sh1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径
      sh2 = sh1[sh1['供应商'].str.contains('海尔施')==False]  # 取反
      sh3 = sh2[~sh2['供应商'].str.contains('恒奇诊断')]  # 取反
      sh4 = sh3[~sh3['供应商'].str.contains('宁波美晶')]  # 取反
      sh5 = sh4[~sh4['供应商'].str.contains('海壹生物')]  # 取反
      sh6 = sh5[~sh5['供应商'].str.contains('强盛生物')]  # 取反
      sh7 = sh6[~sh6['供应商'].str.contains('杭州金诺')]  # 取反

      # sh8 = sh7.loc[sh7.reset_index().groupby(['货品ID'])['成本单价'].idxmax()]  # 取成本单价最高得一笔
      # sh9 = sh8.dropna(how='all')  # 全部缺失就删除

      tkinter.messagebox.showinfo("提醒", "请放入恒奇诊断采购明细单");
      hq1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径
      hq2 = hq1[hq1['供应商'].str.contains('海尔施')==False]  # 取反
      hq3 = hq2[~hq2['供应商'].str.contains('恒奇诊断')]  # 取反
      hq4 = hq3[~hq3['供应商'].str.contains('宁波美晶')]  # 取反
      hq5 = hq4[~hq4['供应商'].str.contains('海壹生物')]  # 取反
      hq6 = hq5[~hq5['供应商'].str.contains('强盛生物')]  # 取反
      hq7 = hq6[~hq6['供应商'].str.contains('杭州金诺')]  # 取反

      tkinter.messagebox.showinfo("提醒", "请放入医药采购明细单");
      yy1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径
      yy2 = yy1[yy1['供应商'].str.contains('海尔施') == False]  # 取反
      yy3 = yy2[~yy2['供应商'].str.contains('恒奇诊断')]  # 取反
      yy4 = yy3[~yy3['供应商'].str.contains('宁波美晶')]  # 取反
      yy5 = yy4[~yy4['供应商'].str.contains('海壹生物')]  # 取反
      yy6 = yy5[~yy5['供应商'].str.contains('强盛生物')]  # 取反
      yy7 = yy6[~yy6['供应商'].str.contains('杭州金诺')]  # 取反




      #####入库明细拼接

      df8 = pd.concat([df5, sh7, hq7,yy7], ignore_index=True)
      df9 = df8.loc[df8.reset_index().groupby(['货品ID'])['成本单价'].idxmax()]  # 取成本单价最高得一笔
      df10 = df9.dropna(how='all')  # 全部缺失就删除


      #####后面自定义销售明细



      tkinter.messagebox.showinfo("提醒", "请放入自定义销售发票明细");
      df5 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      df6 = df5[df5["业务部门"] != '药品一部']
      df7 = df6[df6["业务部门"] != '药品二部']
      df701 = df7.fillna({"客户ID": "无"})
      df702 = df701[df701["客户ID"] != '无']
      df703 = df702.fillna({"业务员": "无"})
      df704 = df703.fillna({"客户所属部门": "无"})
      df705 = df704.fillna({"三级分类": "无"})

      df705['业务员'].astype("str")
      df705['三级分类'].astype("str")
      df705['客户ID'].astype("str")

      # df8 = df705.groupby("客户","客户ID","三级分类","客户所属部门","货品ID")["基本单位数量","价额"].sum();

      df8 = df7.drop(["销售发货总单业务日期", "业务日期", "结算细单ID", "委托人",
                      "通用名","商品名","规格","货品批号","生产日期","有效期","基本单位",
                      "生产厂家","单价","金额","收款标志","累计收款数量","累计收款金额",
                      "未收款金额","税额","税率","总金额","发票号码","发票日期","结算传票ID",
                      "状态","发票类型","外币","汇率","考核日期","应结金额","结算细单备注",
                      "发货总单备注","发货细单备注","制单人","确认人","确认日期","细单条数",
                      "作废标志","作废人","结算总单备注","客户操作码","客户编码","制单人ID",
                      "独立单元ID","货品操作码","业务员ID","业务员操作码","业务部门ID",
                      "结算单ID","预计开票时间", "发票要求", "销售发货细单", "销售发货总单",
                      "保管账ID", "客户部门编号", "存货传票ID","业务部门","产地","独立单元",
                      "未收款数量"
                      ], axis=1)  # 删列

      sw49 = df8[df8["客户"] != '宁波海尔施基因科技有限公司']
      sw4901 = sw49.fillna({"客户": "空"})
      sw4902 = sw4901[sw4901["客户"] != '空']


      sw50 = sw4902.loc[sw4902['客户'].str.contains('海尔施')]
      sw50["关联方"] = '是'
      sw51 = sw4902.loc[sw4902['客户'].str.contains('恒奇诊断')]
      sw51["关联方"] = '是'
      sw52 = sw4902.loc[sw4902['客户'].str.contains('宁波美晶')]
      sw52["关联方"] = '是'
      sw53 = sw4902.loc[sw4902['客户'].str.contains('海壹生物')]
      sw53["关联方"] = '是'
      sw54 = sw4902.loc[sw4902['客户'].str.contains('强盛生物')]
      sw54["关联方"] = '是'
      sw55 = sw4902.loc[sw4902['客户'].str.contains('杭州金诺')]
      sw55["关联方"] = '是'

      sw60 = pd.concat([sw50, sw51, sw52, sw53, sw54, sw55], ignore_index=True)
      # sw601 = d2.drop(["基本单位数量"], axis=1)
      sw61 = pd.merge(sw4902, sw60, how='outer', on=['客户', '货品ID', '三级分类','价额']);
      # sw61["毛利"] = sw61["价额"] - sw61["销售成本"]
      sw62 = sw61.fillna({"关联方": "否"})

      sw63 = sw62[sw62["关联方"] != '是']

      sw63['客户ID_x'].astype(str)
      sw63['价额'].astype("float64")
      sw63['基本单位数量_x'].astype("float64")

      sw64 = sw63.groupby(["客户","客户ID_x","业务员_x","三级分类","客户所属部门_x","货品ID"],as_index=False)["基本单位数量_x","价额"].sum();

      sw65 = sw64[~sw64['三级分类'].str.contains('仪器')]  # 取反
      sw6501 = sw65[~sw65['三级分类'].str.contains('配置')]  # 取反

      # sw65 = sw64[sw64['三级分类'].str.contains('仪器') == False]  # 取反
      # sw66 = sw65[sw65['三级分类'].str.contains('配置') == False]  # 取反

      sw67 = df10.groupby(["货品ID"],as_index=False)["成本单价"].sum();

      sw68 = sw6501.rename(columns={'客户ID_x': '客户ID','业务员_x':'业务员','客户所属部门_x':'客户所属部门','基本单位数量_x':'基本单位数量'});

      sw69 = pd.merge(sw68, sw67, on=['货品ID'],how='left')
      sw70 = sw69.fillna({"成本单价": 0})

      sw70["采购成本"] = (sw70["基本单位数量"].astype("float64") * sw70["成本单价"].astype("float64")) #- sw69["价额"].astype("float64")
      sw71 = sw70.groupby(["客户","客户ID","业务员","客户所属部门"],as_index=False)["价额","采购成本"].sum();





      # ~取反
      # sw66 = sw65[~sw65['三级分类'].str.contains('仪器', na=False)]
      # sw66 = sw65.rename(columns={'三级分类': 'title'})  # 改名
      # sw66 = sw65.loc[~sw4902['三级分类'].str.contains('仪器')]#过期用法

      write = pd.ExcelWriter("本年销售毛利分析" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx")

      sw67.to_excel(write, sheet_name='货品采购成本明细')
      sw68.to_excel(write, sheet_name='货品销售明细')
      sw71.to_excel(write, sheet_name='按客户销售成本明细')

      write.save()
      write.close()
      tkinter.messagebox.showinfo("运行结果", "销售分析都已生成！");






 except Exception as error:
      tm.showerror(title="煎饼提示前方路堵",
              message="请检查提交源文件是否正确 '" + str(error) + "'.",
               detail=traceback.format_exc())



############################
def appendStr100():  #测试组
 try:
      tkinter.messagebox.showinfo("提醒当前版本2.5", "下载地址：待定");
      tkinter.messagebox.showinfo("提醒", "请选择生物源文件");
      df1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      df2 = df1.groupby(['客户', '货品ID', '三级分类'])["销售成本", "基本单位数量", "价额"].sum();
      df2.reset_index(inplace=True)  # 取消合并

      d2 = df1.groupby(['客户', '货品ID', '三级分类'])["销售成本", "基本单位数量", "价额"].sum();
      d2.reset_index(inplace=True)  # 取消合并
      tkinter.messagebox.showinfo("提醒", "请选择分类规则源文件");
      df63 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

      sw49 = d2[d2["客户"] != '宁波海尔施基因科技有限公司']
      sw50 = sw49.loc[sw49['客户'].str.contains('海尔施')]
      sw50["关联方"] = '是'
      sw51 = sw49.loc[sw49['客户'].str.contains('恒奇诊断')]
      sw51["关联方"] = '是'
      sw52 = sw49.loc[sw49['客户'].str.contains('宁波美晶')]
      sw52["关联方"] = '是'
      sw53 = sw49.loc[sw49['客户'].str.contains('海壹生物')]
      sw53["关联方"] = '是'
      sw54 = sw49.loc[sw49['客户'].str.contains('强盛生物')]
      sw54["关联方"] = '是'
      sw55 = sw49.loc[sw49['客户'].str.contains('杭州金诺')]
      sw55["关联方"] = '是'

      sw60 = pd.concat([sw50, sw51, sw52, sw53, sw54, sw55], ignore_index=True)
      # sw601 = d2.drop(["基本单位数量"], axis=1)
      sw61 = pd.merge(d2, sw60, how='outer', on=['客户', '货品ID', '三级分类', '销售成本', '价额', '基本单位数量']);
      sw61["毛利"] = sw61["价额"] - sw61["销售成本"]
      sw62 = sw61.fillna({"关联方": "否"})
      sw63 = sw62.groupby(['三级分类'])["销售成本", "价额", "毛利"].sum();
      sw64 = pd.merge(sw62, df63, how='left', on=['三级分类']);  # 63是输入
      sw65 = sw64.groupby(['大类'])["销售成本", "价额", "毛利"].sum();

      write = pd.ExcelWriter("毛利分析1" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx")

      sw62.to_excel(write, sheet_name='单体明细')
      sw63.to_excel(write, sheet_name='三级分类')
      sw65.to_excel(write, sheet_name='大类')
      write.save()
      write.close()

      #################################################################11-17
      # df1 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径
      #
      # df2 = df1.loc[df1.reset_index().groupby(['货品ID'])['成本单价'].idxmax()]      #取成本单价最高得一笔
      # df3 = df2.dropna(how='all')  #全部缺失就删除
      #
      # tkinter.messagebox.showinfo("提醒", "请放入自定义销售发票明细");
      # df5 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径
      #
      # df6 = df5[df5["业务部门"] != '药品一部']
      # df7 = df6[df6["业务部门"] != '药品二部']
      # df701 = df7.fillna({"客户ID": "无"})
      # df702 = df701[df701["客户ID"] != '无']
      # df703 = df702.fillna({"业务员": "无"})
      # df704 = df703.fillna({"客户所属部门": "无"})
      # df705 = df704.fillna({"三级分类": "无"})
      #
      # df705['业务员'].astype("str")
      # df705['三级分类'].astype("str")
      # df705['客户ID'].astype("str")
      #
      #
      # df8 = df7.drop(["销售发货总单业务日期", "业务日期", "结算细单ID", "委托人",
      #                 "通用名","商品名","规格","货品批号","生产日期","有效期","基本单位",
      #                 "生产厂家","单价","金额","收款标志","累计收款数量","累计收款金额",
      #                 "未收款金额","税额","税率","总金额","发票号码","发票日期","结算传票ID",
      #                 "状态","发票类型","外币","汇率","考核日期","应结金额","结算细单备注",
      #                 "发货总单备注","发货细单备注","制单人","确认人","确认日期","细单条数",
      #                 "作废标志","作废人","结算总单备注","客户操作码","客户编码","制单人ID",
      #                 "独立单元ID","货品操作码","业务员ID","业务员操作码","业务部门ID",
      #                 "结算单ID","预计开票时间", "发票要求", "销售发货细单", "销售发货总单",
      #                 "保管账ID", "客户部门编号", "存货传票ID","业务部门","产地","独立单元",
      #                 "未收款数量"
      #                 ], axis=1)  # 删列
      #
      # sw49 = df8[df8["客户"] != '宁波海尔施基因科技有限公司']
      # sw4901 = sw49.fillna({"客户": "空"})
      # sw4902 = sw4901[sw4901["客户"] != '空']
      #
      #
      # sw50 = sw4902.loc[sw4902['客户'].str.contains('海尔施')]
      # sw50["关联方"] = '是'
      # sw51 = sw4902.loc[sw4902['客户'].str.contains('恒奇诊断')]
      # sw51["关联方"] = '是'
      # sw52 = sw4902.loc[sw4902['客户'].str.contains('宁波美晶')]
      # sw52["关联方"] = '是'
      # sw53 = sw4902.loc[sw4902['客户'].str.contains('海壹生物')]
      # sw53["关联方"] = '是'
      # sw54 = sw4902.loc[sw4902['客户'].str.contains('强盛生物')]
      # sw54["关联方"] = '是'
      # sw55 = sw4902.loc[sw4902['客户'].str.contains('杭州金诺')]
      # sw55["关联方"] = '是'
      #
      # sw60 = pd.concat([sw50, sw51, sw52, sw53, sw54, sw55], ignore_index=True)
      #
      # sw61 = pd.merge(sw4902, sw60, how='outer', on=['客户', '货品ID', '三级分类','价额']);
      #
      # sw62 = sw61.fillna({"关联方": "否"})
      #
      # sw63 = sw62[sw62["关联方"] != '是']
      #
      # sw63['客户ID_x'].astype(str)
      # sw63['价额'].astype("float64")
      # sw63['基本单位数量_x'].astype("float64")
      #
      # sw64 = sw63.groupby(["客户","客户ID_x","业务员_x","三级分类","客户所属部门_x","货品ID"],as_index=False)["基本单位数量_x","价额"].sum();
      #
      # sw65 = sw64[~sw64['三级分类'].str.contains('仪器')]  # 取反
      #
      #
      #
      # sw67 = df3.groupby(["货品ID"],as_index=False)["成本单价"].sum();
      #
      # sw68 = sw65.rename(columns={'客户ID_x': '客户ID','业务员_x':'业务员','客户所属部门_x':'客户所属部门','基本单位数量_x':'基本单位数量'});
      #
      # sw69 = pd.merge(sw68, sw67, on=['货品ID'])
      # sw69["毛利"] = (sw69["基本单位数量"].astype("float64") * sw69["成本单价"].astype("float64")) - sw69["价额"].astype("float64")
      #
      # sw70 = sw69.groupby(["客户","客户ID","业务员","客户所属部门"],as_index=False)["价额","基本单位数量","成本单价"].sum();
      #
      #
      #
      # write = pd.ExcelWriter("本年销售毛利分析" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx")
      #
      # sw67.to_excel(write, sheet_name='采购成本明细')
      #
      # sw70.to_excel(write, sheet_name='按客户销售毛利明细')
      #
      # write.save()
      # write.close()
      # tkinter.messagebox.showinfo("运行结果", "单体毛利分析都已生成！");

     ###################################################



      # df3.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
      #                                                                  filetypes=[("Microsoft Excel文件", "*.xlsx"),
      #                                                                             (
      #                                                                                "Microsoft Excel 97-20003 文件",
      #                                                                                "*.xls")],
      #                                                                  defaultextension=".xlsx"), sheet_name="表1");
      # tkinter.messagebox.showinfo("运行结果", "采购数据导出成功！");
      #
      # sw69.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
      #                                                                filetypes=[("Microsoft Excel文件", "*.xlsx"),
      #                                                                           (
      #                                                                              "Microsoft Excel 97-20003 文件",
      #                                                                              "*.xls")],
      #                                                                defaultextension=".xlsx"), sheet_name="表1");
      # tkinter.messagebox.showinfo("运行结果", "销售明细处理导出成功！");



      # sw62 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径
      # df63 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径
      # sw63 = sw62.groupby(['三级分类'])["销售成本", "价额", "毛利"].sum();
      #
      #
      #
      # sw64 = pd.merge(sw62, df63, how='left', on=['三级分类']);  # 63是输入
      # sw65 = sw64.groupby(['大类'])["销售成本", "价额", "毛利"].sum();
      #
      #
      #
      # write = pd.ExcelWriter("生物毛利分析" + str(datetime.datetime.now().strftime('%Y%m%d')) + ".xlsx")
      #
      # sw63.to_excel(write, sheet_name='三级分类')
      # sw65.to_excel(write, sheet_name='大类')
      # write.save()
      # write.close()
      # tkinter.messagebox.showinfo("运行结果", "完成！");

 except Exception as error:
      tm.showerror(title="煎饼提示前方路堵",
              message="请检查提交源文件是否正确 '" + str(error) + "'.",
               detail=traceback.format_exc())

##############
############################
# def appendStr99():  #测试组
#  try:
#     os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
#
#     # 查询操作（查）
#
#     db = oracle.connect('heseas/kingdee@60.12.218.220:6694/ORCLEAS')  # 数据库连接
#     cursor = db.cursor()  # 创建cursor
#     cursor.execute("SELECT * FROM t_gl_acctcussent")  # 执行sql语句
#     rs = cursor.fetchall()  # 一次返回所有结果集 fetchall
#     id2 = rs[0][0]  # 去除多余的内容
#     print(id2)  # 打印内容
#
#     db.close()  # 关闭数据库连接
#
#     tkinter.messagebox.showinfo(id2, "下载地址：待定");
#
#
#  except Exception as error:
#       tm.showerror(title="煎饼提示前方路堵",
#               message="请检查提交源文件是否正确 '" + str(error) + "'.",
#                detail=traceback.format_exc())

##############
filemenu=Menu(menubar,tearoff=False)

# filemenu.add_checkbutton(label="客户汇总",command=appendStr)
# filemenu.add_checkbutton(label="发票汇总",command=appendStr2)
filemenu.add_command(label="客户汇总",command=appendStr)
filemenu.add_command(label="发票汇总",command=appendStr2)
filemenu.add_command(label="税率明细",command=appendStr1)
filemenu.add_command(label="简易征收明细",command=appendStr20)
filemenu.add_command(label="客户+税率+票号",command=appendStr81)
filemenu.add_command(label="系统退出",command=root.quit)




menubar.add_cascade(label='发票功能',menu=filemenu)


file2menu=Menu(menubar,tearoff=False)

file2menu.add_command(label="计算发出商品",command=appendStr5)
file2menu.add_command(label="统计未开票客户",command=appendStr6)
file2menu.add_command(label="关联交易货品分类",command=appendStr14)
file2menu.add_command(label="货品成本分析",command=appendStr17)
file2menu.add_command(label="客户成本分析",command=appendStr18)

menubar.add_cascade(label='英克功能',menu=file2menu)








file3menu=Menu(menubar,tearoff=False)

file3menu.add_command(label="单体账龄分列",command=appendStr3)
file3menu.add_command(label="不带项目账龄分列",command=appendStr41)
file3menu.add_command(label="辅助总账分列",command=appendStr8)
file3menu.add_command(label="集团账龄分列",command=appendStr9)
file3menu.add_command(label="医药账龄计算逾期",command=appendStr10)
file3menu.add_command(label="医药回款考核",command=appendStr28)
file3menu.add_command(label="对账单按客户整理",command=appendStr16)
file3menu.add_command(label="非公控制整理表",command=appendStr19)
file3menu.add_command(label="非公控制整理表-带英克ID",command=appendStr84)
file3menu.add_command(label="销售英克-金蝶核对",command=appendStr83)
file3menu.add_command(label="暂估英克-金蝶核对",command=appendStr86)

menubar.add_cascade(label='金蝶功能',menu=file3menu)



file4menu=Menu(menubar,tearoff=False)
file4menu.add_command(label="发票代码+号码+日期",command=appendStr101)
file4menu.add_command(label="发票代码+号码",command=appendStr102)
file4menu.add_command(label="发票号码",command=appendStr103)



menubar.add_cascade(label='认证功能',menu=file4menu)






file6menu=Menu(menubar,tearoff=False)
file6menu.add_command(label="2020考核表",command=appendStr32)
file6menu.add_command(label="2020英克毛利计算",command=appendStr101)
# file6menu.add_command(label="历史销售总单",command=appendStr4)
# file6menu.add_command(label="英克销售与开票对比",command=appendStr21)
# file6menu.add_command(label="英克出库汇总",command=appendStr22)
# file6menu.add_command(label="英克出库招标价对比",command=appendStr24)
# file6menu.add_command(label="英克15-19年销售分析",command=appendStr25)
# file6menu.add_command(label="销售测算合并仪器",command=appendStr26)
# # file6menu.add_command(label="本年仪器回顾",command=appendStr29)
# file6menu.add_command(label="英克年度销售分析",command=appendStr31)

menubar.add_cascade(label='数据测算',menu=file6menu)


file7menu=Menu(menubar,tearoff=False)
file7menu.add_command(label="往来抵消行列转换",command=appendStr82)
file7menu.add_command(label="DHY-K3转换",command=appendStr43)
file7menu.add_command(label="QS-K3转换",command=appendStr44)
file7menu.add_command(label="JCK-K3转换",command=appendStr45)
file7menu.add_command(label="非关联交易毛利计算-包含单体",command=appendStr85)

menubar.add_cascade(label='报表数据',menu=file7menu)



file5menu=Menu(menubar,tearoff=False)
file5menu.add_command(label="网页验真独包如需要请联系煎饼")
file5menu.add_command(label="图片识别独包如需要请联系煎饼")

# file5menu.add_command(label="测试功能",command=appendStr99)

menubar.add_cascade(label='识别验真',menu=file5menu)

file8menu=Menu(menubar,tearoff=False)
file8menu.add_command(label="2.0新界面")
file8menu.add_command(label="遇到错误请拍砖")
file8menu.add_command(label="新需求请联系煎饼")
file8menu.add_command(label="下载新版",command=appendStr100)
menubar.add_cascade(label='版本介绍',menu=file8menu)

root.config(menu=menubar)
mainloop()











# # -*- coding: UTF-8 -*-
# # Python连接Oracle数据库实现增删改查
#
# import cx_Oracle as oracle
# import os
#
# os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
#
#
# # 查询操作（查）
#
#
#
# db=oracle.connect('heseas/kingdee@60.12.218.220:6694/orcleas')#数据库连接
# cursor=db.cursor()#创建cursor
# cursor.execute("SELECT * FROM t_gl_acctcussent")#执行sql语句
# rs=cursor.fetchall()#一次返回所有结果集 fetchall
# id2=rs[0][0]#去除多余的内容
# print(id2)#打印内容
#
# db.close()#关闭数据库连接
#

df156 = pd.merge(df124, df155, how='left', on=['客户名称']);  # 完全相同合并，忽略没有的客户(没有how)

    #####自动生成完全版数据

    #df157 = pd.merge(df123, df155, how='left', on=['客户名称']);

    #df151 = df150['项目代码'].str.split('-| ', expand=True);

    #df152 = pd.merge(df151, df23, right_index=True, left_index=True);

    df261 = df25[df25["公司"] != "海尔施集团"]
    df262 = df261[df261["项目"] == "001 代理试剂"]

    ####19-10-25整理输出表格格式

    #df157 = df156.drop(["Unnamed: 6_x", "等级_y", "Unnamed: 6_y"], axis=1)

    #df158 = df157.rename(columns={'等级_x': '等级'});



    df156.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               ("Microsoft Excel 97-20003 文件",
                                                                                   "*.xls")],
                                                                    defaultextension=".xlsx"));

   # df262.to_excel("客户等级全部公司版" + str(datetime.datetime.now().strftime('%Y%m%d%h')) + ".xls", sheet_name="sheet1",index=False)  # 自动输出
    tkinter.messagebox.showinfo("运行结果", "客户等级测试导出成功！");

 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
             message="请检查提交源文件是否正确 '" + str(error) + "'.",
             detail=traceback.format_exc())




def appendStr20():  ##########################################################################################销售发票分税率巧儿

 try:
    tkinter.messagebox.showinfo("提醒", "请先选择开票明细源文件");
    df = pd.read_excel(tkinter.filedialog.askopenfilename());

    df1 = df.drop(df.columns[[[[[[[[[[3, 4, 5, 7, 8, 10, 11, 12, 13, 17]]]]]]]]]], axis=1);
    df2 = df1.drop(df1.index[[[[0, 1, 2, 3]]]], axis=0);
    df3 = df2[df2["Unnamed: 9"] != "小计"];
    df4 = df3[df3["Unnamed: 9"] != "商品名称"];
    df5 = df4.dropna(how="all");
    df6 = df5.fillna(method='pad');
    df6["Unnamed: 14"] = df6["Unnamed: 14"].astype("float64");  # 改变格式
    df6["Unnamed: 16"] = df6["Unnamed: 16"].astype("float64");

    df9 = df6.loc[df6['Unnamed: 9'].str.contains('紫杉醇注射液')]  # 7-15模糊查询,但单元格不能为0为空  （ #print(df4[df4['客户'].isin(['宁波海尔施医学检验所有限公司'])])）
    df10 = df6.loc[df6['Unnamed: 9'].str.contains('阿那曲唑片')]
    df11 = df6.loc[df6['Unnamed: 9'].str.contains('奥沙利铂甘露醇注射液')]
    df12 = df6.loc[df6['Unnamed: 9'].str.contains('比卡鲁胺片')]
    df13 = df6.loc[df6['Unnamed: 9'].str.contains('醋酸奥曲肽注射液')]
    df14 = df6.loc[df6['Unnamed: 9'].str.contains('醋酸戈舍瑞林缓释植入剂')]
    df15 = df6.loc[df6['Unnamed: 9'].str.contains('多西他赛注射液')]
    df16 = df6.loc[df6['Unnamed: 9'].str.contains('吉非替尼片')]
    df17 = df6.loc[df6['Unnamed: 9'].str.contains('甲苯磺酸索拉非尼片')]
    df18 = df6.loc[df6['Unnamed: 9'].str.contains('甲磺酸奥希替尼片')]
    df19 = df6.loc[df6['Unnamed: 9'].str.contains('甲磺酸伊马替尼片')]
    df20 = df6.loc[df6['Unnamed: 9'].str.contains('酒石酸长春瑞滨注射液')]
    df21 = df6.loc[df6['Unnamed: 9'].str.contains('卡培他滨片')]
    df22 = df6.loc[df6['Unnamed: 9'].str.contains('来曲唑片')]
    df23 = df6.loc[df6['Unnamed: 9'].str.contains('硫培非格司亭注射液')]
    df24 = df6.loc[df6['Unnamed: 9'].str.contains('顺铂注射液')]
    df25 = df6.loc[df6['Unnamed: 9'].str.contains('替吉奥胶囊')]
    df26 = df6.loc[df6['Unnamed: 9'].str.contains('注射用奥沙利铂')]
    df27 = df6.loc[df6['Unnamed: 9'].str.contains('注射用地西他滨')]
    df28 = df6.loc[df6['Unnamed: 9'].str.contains('注射用洛铂')]
    df29 = df6.loc[df6['Unnamed: 9'].str.contains('注射用奈达铂')]
    df30 = df6.loc[df6['Unnamed: 9'].str.contains('注射用培美曲塞二钠')]
    df31 = df6.loc[df6['Unnamed: 9'].str.contains('注射用亚叶酸钙')]
    df32 = df6.loc[df6['Unnamed: 9'].str.contains('注射用盐酸表柔比星')]
    df33 = df6.loc[df6['Unnamed: 9'].str.contains('注射用盐酸吉西他滨')]
    df34 = df6.loc[df6['Unnamed: 9'].str.contains('注射用盐酸伊立替康')]
    df35 = df6.loc[df6['Unnamed: 9'].str.contains('注射用紫杉醇')]
    df36 = df6.loc[df6['Unnamed: 9'].str.contains('比卡鲁胺胶囊')]

    df40 = pd.concat([df9, df10, df11, df12, df13, df14, df15, df16, df17, df18, df19, df20, df21, df22,df23,
                      df24,df25,df26,df27,df28,df29,df30,df31,df32,df33,df34,df35,df36],
                      ignore_index=True)  # 组合
    df41 = df40.sort_values(by=['Unnamed: 1'], axis=0, ascending=True)  # 行排序

    df42 = df41.rename(columns={'Unnamed: 1': '发票号码', 'Unnamed: 2': '客户', 'Unnamed: 6': '发票日期', 'Unnamed: 9': '货品名称',
                         'Unnamed: 14': '无税金额', 'Unnamed: 15': '税率','Unnamed: 16': '税额'});


    df42["分类"]=df42["税率"]
    df42["分类"].replace("3%", "抗癌3%", inplace=True)




    #######上面是抗癌药物3%
    #####下面开始药品3%
    df49 = df6[df6["Unnamed: 15"] == "3%"];

    df50 = df49.loc[df49['Unnamed: 9'].str.contains('重组人干扰素a2b注射液')]  # 7-15模糊查询,但单元格不能为0为空
    df51 = df49.loc[df49['Unnamed: 9'].str.contains('注射用鼠神经生长因子')]  # 7-15模糊查询,但单元格不能为0为空
    df52 = df49.loc[df49['Unnamed: 9'].str.contains('脑苷肌肽注射液')]  # 7-15模糊查询,但单元格不能为0为空
    df53 = df49.loc[df49['Unnamed: 9'].str.contains('骨瓜提取物注射液')]  # 7-15模糊查询,但单元格不能为0为空
    df54 = df49.loc[df49['Unnamed: 9'].str.contains('注射用骨肽')]  # 7-15模糊查询,但单元格不能为0为空
    df55 = df49.loc[df49['Unnamed: 9'].str.contains('静注人免疫球蛋白')]  # 7-15模糊查询,但单元格不能为0为空
    df56 = df49.loc[df49['Unnamed: 9'].str.contains('人血白蛋白')]  # 7-15模糊查询,但单元格不能为0为空
    df57 = df49.loc[df49['Unnamed: 9'].str.contains('人凝血酶原复合物')]  # 7-15模糊查询,但单元格不能为0为空
    df58 = df49.loc[df49['Unnamed: 9'].str.contains('破伤风人免疫球蛋白')]  # 7-15模糊查询,但单元格不能为0为空
    df59 = df49.loc[df49['Unnamed: 9'].str.contains('缩宫素注射液')]  # 7-15模糊查询,但单元格不能为0为空
    df60 = df49.loc[df49['Unnamed: 9'].str.contains('注射用硼替佐米')]  # 7-15模糊查询,但单元格不能为0为空
    df61 = df49.loc[df49['Unnamed: 9'].str.contains('注射用白眉蛇毒血凝酶')]  # 7-15模糊查询,但单元格不能为0为空
    df62 = df49.loc[df49['Unnamed: 9'].str.contains('酪酸梭菌活菌胶囊')]  # 7-15模糊查询,但单元格不能为0为空

    df63 = pd.concat([df50,df51,df52,df53,df54,df55,df56,df57,df58,df59,df60,df61,df62],ignore_index=True)  # 组合
    print(df63)

    df64 = df63.sort_values(by=['Unnamed: 1'], axis=0, ascending=True)  # 行排序

    df65 = df64.rename(columns={'Unnamed: 1': '发票号码', 'Unnamed: 2': '客户', 'Unnamed: 6': '发票日期', 'Unnamed: 9': '货品名称',
                                'Unnamed: 14': '无税金额', 'Unnamed: 15': '税率', 'Unnamed: 16': '税额'});

    df65["分类"] = df65["税率"]
    df65["分类"].replace("3%", "药品3%", inplace=True)



    #####药品3%结束
    #####试剂3%开始
    df66 = df6[df6["Unnamed: 15"] == "3%"];

    df67 = df66.loc[df66['Unnamed: 9'].str.contains('A抗A抗B血型定型试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df68 = df66.loc[df66['Unnamed: 9'].str.contains('B抗A抗B血型定型试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df69 = df66.loc[df66['Unnamed: 9'].str.contains('抗A,抗B血型定型试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df70 = df66.loc[df66['Unnamed: 9'].str.contains('0605005人类免疫缺陷病毒抗体诊断试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df71 = df66.loc[df66['Unnamed: 9'].str.contains('0605007乙型肝炎病毒表面抗原诊断试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df72 = df66.loc[df66['Unnamed: 9'].str.contains('A 抗A抗B血型定型试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df73 = df66.loc[df66['Unnamed: 9'].str.contains('B 抗A抗B血型定型试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df74 = df66.loc[df66['Unnamed: 9'].str.contains('抗人球蛋白检测卡')]  # 7-15模糊查询,但单元格不能为0为空
    df75 = df66.loc[df66['Unnamed: 9'].str.contains('梅毒螺旋体抗体诊断试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df76 = df66.loc[df66['Unnamed: 9'].str.contains('19211人类免疫缺陷病毒抗体诊断试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df77 = df66.loc[df66['Unnamed: 9'].str.contains('乙型肝炎病毒核心抗体IgM检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df78 = df66.loc[df66['Unnamed: 9'].str.contains('乙型肝炎病毒前S1抗原检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df79 = df66.loc[df66['Unnamed: 9'].str.contains('丙型肝炎病毒抗体诊断试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df80 = df66.loc[df66['Unnamed: 9'].str.contains('梅毒甲苯胺红不加热血清试验诊断试剂')]  # 7-15模糊查询,但单元格不能为0为空
    df81 = df66.loc[df66['Unnamed: 9'].str.contains('ABO血型反定型试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df82 = df66.loc[df66['Unnamed: 9'].str.contains('ABO、RhD血型定型检测卡')]  # 7-15模糊查询,但单元格不能为0为空
    df83 = df66.loc[df66['Unnamed: 9'].str.contains('甲型肝炎病毒IgM抗体检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df84 = df66.loc[df66['Unnamed: 9'].str.contains('A乙型肝炎病毒表面抗体检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df85 = df66.loc[df66['Unnamed: 9'].str.contains('01200205A乙型肝炎病毒e抗原检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df86 = df66.loc[df66['Unnamed: 9'].str.contains('01200208A乙型肝炎病毒核心抗体检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空
    df87 = df66.loc[df66['Unnamed: 9'].str.contains('01200210A乙型肝炎病毒e抗体检测试剂盒')]  # 7-15模糊查询,但单元格不能为0为空

    df95 = pd.concat([df67,df68,df69,df70,df71,df72,df73,df74,df75,df76,df77,df78,df79,df80,df81,df82,df83,df84,df85,df86,df87],
                     ignore_index=True)  # 组合
    print(df95)

    df96 = df95.sort_values(by=['Unnamed: 1'], axis=0, ascending=True)  # 行排序

    df97 = df96.rename(columns={'Unnamed: 1': '发票号码', 'Unnamed: 2': '客户', 'Unnamed: 6': '发票日期', 'Unnamed: 9': '货品名称',
                                'Unnamed: 14': '无税金额', 'Unnamed: 15': '税率', 'Unnamed: 16': '税额'});

    df97["分类"] = df97["税率"]
    df97["分类"].replace("3%", "试剂3%", inplace=True)

   ####试剂3%完毕
   ####其他3%

    df110 = df6[df6["Unnamed: 15"] == "3%"];

    A=df110['Unnamed: 9'].str.contains('A抗A抗B血型定型试剂')

    print(A)
    df111 = df110[df110["Unnamed: 9"] != A]




    df112= df111.sort_values(by=['Unnamed: 1'], axis=0, ascending=True)  # 行排序

    df113 = df112.rename(columns={'Unnamed: 1': '发票号码', 'Unnamed: 2': '客户', 'Unnamed: 6': '发票日期', 'Unnamed: 9': '货品名称',
                                'Unnamed: 14': '无税金额', 'Unnamed: 15': '税率', 'Unnamed: 16': '税额'});

    df100 =pd.concat([df42,df65,df97,df113],
                     ignore_index=True)  # 组合

    df7a= df6.rename(columns={'Unnamed: 1': '发票号码', 'Unnamed: 2': '客户', 'Unnamed: 6': '发票日期', 'Unnamed: 9': '货品名称',
                                'Unnamed: 14': '无税金额', 'Unnamed: 15': '税率', 'Unnamed: 16': '税额'});
    df102 = pd.concat([df100,df7a],ignore_index=True)  # 组合
    # 查看是否有重复行
    re_row = df100.duplicated()
    print(re_row)

    # 查看去除重复行的数据
    no_re_row = df100.drop_duplicates()
    print(no_re_row)

    # 查看基于[物品]列去除重复行的数据f
    df101 = df102.drop_duplicates(['货品名称','无税金额','客户','发票号码'])
    #print(wp)
    df101["金额"]=df101["无税金额"]+df101["税额"]
    df101.to_excel(excel_writer=tkinter.filedialog.asksaveasfilename(title="请创建或者选择一个保存数据的Excel文件",
                                                                    filetypes=[("Microsoft Excel文件", "*.xlsx"),
                                                                               (
                                                                                   "Microsoft Excel 97-20003 文件",
                                                                                   "*.xls")],
                                                                    defaultextension=".xlsx"));


    #df101.to_excel(excel_writer="d:/英克测试.xlsx",
     #             sheet_name="测试1",
      #            );
    tkinter.messagebox.showinfo("运行结果", "开票税率分类导出成功！");
 except Exception as error:
     tm.showerror(title="煎饼提示前方路堵",
             message="请检查提交源文件是否正确 '" + str(error) + "'.",
             detail=traceback.format_exc())





def appendStr21():  #######英克销售和开票对比

 try:
    tkinter.messagebox.showinfo("提醒", "请先选择出库明细源文件");
    df1 = pd.read_excel(tkinter.filedialog.askopenfilename());
    df2=df1.fillna(0)
    df3 = df2.drop(df2.index[[0, 1]], axis=0);

    df3["仪器出库合计"]=df3["03原厂仪器"]+df3["04采购平台仪器"]+df3["0501国产辅助配置"]+df3["0502流水线辅助配置"]

    df3["非仪器出库合计"]=df3["010101免疫（代理）"]+df3["010102特定蛋白（代理）"]+df3["010103血球（代理）"]+df3["010104普通生化（代理）"]\
    +df3["010105AU生化（代理）"]+df3["010106利德曼生化（代理）"]+df3["010107尿液（代理）"]+df3["010109微生物（代理）"]+df3["010110索灵（代理BC）"]\
    +df3["010111免疫（AMH）"]+df3["010201血凝（代理）"]+df3["0103lmmucor"]+df3["0104索灵"]+df3["010501质控试剂（代理）"]+df3["010502伯乐其它试剂"]\
    +df3["010601BNP试剂（代理）"]+df3["010701血气（代理）"]+df3["0108苏医（代理BC血球质控）"]+df3["020101干式生化"]+df3["020102普通生化"]\
    +df3["020103血气"]+df3["020104特殊生化"]+df3["020201血球"]+df3["020202血凝"]+df3["020203尿液"]+df3["020204血库"]+df3["020206体液"]\
    +df3["020301发光"] +df3["020302特定蛋白"]+df3["020303酶免类"]+df3["020304其它免疫"]+df3["020305厦门万泰"]+df3["0204微生物"]+df3["0205药字号"]\
    +df3["0206分子诊断"]+df3["0207病理科"]+df3["0208采购平台其它"]+df3["0209质控"]+df3["06软件"]+df3["07配件"] \
    +df3["08其它业务"]+df3["0901基因试剂（自产）"]+df3["0902基因试剂（其它厂家）"]+df3["1101强盛生化"]+df3["1201沃文特免疫"]+df3["1202沃文特其他"]\
    +df3["99其它"]

    df4 = df3.drop(["Unnamed: 1", "Unnamed: 2", "Unnamed: 3"], axis=1)  # 删列

    df5 = df4.groupby(["Unnamed: 0"], as_index=False)["非仪器出库合计", "仪器出库合计"].sum();

    #df5["Unnamed: 0"]=df5[""]
    df5["地区"] = df5["Unnamed: 0"]
    df5["负责人"] = df5["Unnamed: 0"]
    df5["地区编码"] = df5["Unnamed: 0"]
    df5["部门"] = df5["Unnamed: 0"]


    df5["地区"].replace("温州葛瑞", "温州1", inplace=True)
    df5["地区编码"].replace("温州葛瑞", "0101", inplace=True)
    df5["负责人"].replace("温州葛瑞", "葛瑞", inplace=True)
    df5["部门"].replace("温州葛瑞", "01部", inplace=True)

    df5["地区"].replace("台州唐惠", "台州1", inplace=True)
    df5["地区编码"].replace("台州唐惠", "0103", inplace=True)
    df5["负责人"].replace("台州唐惠", "唐惠", inplace=True)
    df5["部门"].replace("台州唐惠", "01部", inplace=True)

    df5["地区"].replace("温州潘磊", "温州2", inplace=True)
    df5["地区编码"].replace("温州潘磊", "0102", inplace=True)
    df5["负责人"].replace("温州潘磊", "潘磊", inplace=True)
    df5["部门"].replace("温州潘磊", "01部", inplace=True)

    df5["地区"].replace("台州胡文魁", "台州2", inplace=True)
    df5["地区编码"].replace("台州胡文魁", "0104", inplace=True)
    df5["负责人"].replace("台州胡文魁", "胡文魁", inplace=True)
    df5["部门"].replace("台州胡文魁", "01部", inplace=True)

    df5["地区"].replace("丽水", "丽水", inplace=True)
    df5["地区编码"].replace("丽水", "0105", inplace=True)
    df5["负责人"].replace("丽水", "方汝泼", inplace=True)
    df5["部门"].replace("丽水", "01部", inplace=True)


   #####一部完毕
    df5["地区"].replace("宁波市区", "宁波", inplace=True)
    df5["地区编码"].replace("宁波市区", "0201", inplace=True)
    df5["负责人"].replace("宁波市区", "丁玲", inplace=True)
    df5["部门"].replace("宁波市区", "02部", inplace=True)

    df5["地区"].replace("舟山北仑", "舟山北仑", inplace=True)
    df5["地区编码"].replace("舟山北仑", "0202", inplace=True)
    df5["负责人"].replace("舟山北仑", "高大勇", inplace=True)
    df5["部门"].replace("舟山北仑", "02部", inplace=True)

    df5["地区"].replace("慈溪余姚镇海", "北三县", inplace=True)
    df5["地区编码"].replace("慈溪余姚镇海", "0203", inplace=True)
    df5["负责人"].replace("慈溪余姚镇海", "陆金耀", inplace=True)
    df5["部门"].replace("慈溪余姚镇海", "02部", inplace=True)

    df5["地区"].replace("奉化宁海象山", "南三县", inplace=True)
    df5["地区编码"].replace("奉化宁海象山", "0204", inplace=True)
    df5["负责人"].replace("奉化宁海象山", "吴燕江", inplace=True)
    df5["部门"].replace("奉化宁海象山", "02部", inplace=True)
   ####三部####

    df5["地区"].replace("杭州姜立民", "省级", inplace=True)
    df5["地区编码"].replace("杭州姜立民", "0301", inplace=True)
    df5["负责人"].replace("杭州姜立民", "姜立民", inplace=True)
    df5["部门"].replace("杭州姜立民", "03部", inplace=True)

    df5["地区"].replace("杭州石亚国", "省级", inplace=True)
    df5["地区编码"].replace("杭州石亚国", "0301", inplace=True)
    df5["负责人"].replace("杭州石亚国", "姜立民", inplace=True)
    df5["部门"].replace("杭州石亚国", "03部", inplace=True)

    df5["地区"].replace("杭州陈靓", "省级", inplace=True)
    df5["地区编码"].replace("杭州陈靓", "0301", inplace=True)
    df5["负责人"].replace("杭州陈靓", "姜立民", inplace=True)
    df5["部门"].replace("杭州陈靓", "03部", inplace=True)

    df5["地区"].replace("杭州沈剑芳", "杭州", inplace=True)
    df5["地区编码"].replace("杭州沈剑芳", "0302", inplace=True)
    df5["负责人"].replace("杭州沈剑芳", "沈剑芳", inplace=True)
    df5["部门"].replace("杭州沈剑芳", "03部", inplace=True)

    df5["地区"].replace("杭州周海波", "杭州", inplace=True)
    df5["地区编码"].replace("杭州周海波", "0302", inplace=True)
    df5["负责人"].replace("杭州周海波", "沈剑芳", inplace=True)
    df5["部门"].replace("杭州周海波", "03部", inplace=True)

    df5["地区"].replace("嘉兴阮芳", "嘉湖", inplace=True)
    df5["地区编码"].replace("嘉兴阮芳", "0303", inplace=True)
    df5["负责人"].replace("嘉兴阮芳", "阮芳", inplace=True)
    df5["部门"].replace("嘉兴阮芳", "03部", inplace=True)

    df5["地区"].replace("湖州陈荣斌", "嘉湖", inplace=True)
    df5["地区编码"].replace("湖州陈荣斌", "0303", inplace=True)
    df5["负责人"].replace("湖州陈荣斌", "阮芳", inplace=True)
    df5["部门"].replace("湖州陈荣斌", "03部", inplace=True)

    df5["地区"].replace("晋江运城郑良", "晋江运城", inplace=True)
    df5["地区编码"].replace("晋江运城郑良", "0304", inplace=True)
    df5["负责人"].replace("晋江运城郑良", "郑良", inplace=True)
    df5["部门"].replace("晋江运城郑良", "03部", inplace=True)
  ####四部

    df5["地区"].replace("南京高跃", "南京1", inplace=True)
    df5["地区编码"].replace("南京高跃", "0401", inplace=True)
    df5["负责人"].replace("南京高跃", "高跃", inplace=True)
    df5["部门"].replace("南京高跃", "04部", inplace=True)

    df5["地区"].replace("南京阮建锋", "南京2", inplace=True)
    df5["地区编码"].replace("南京阮建锋", "0402", inplace=True)
    df5["负责人"].replace("南京阮建锋", "阮建锋", inplace=True)
    df5["部门"].replace("南京阮建锋", "04部", inplace=True)

    df5["地区"].replace("南京刘纪彬", "南京3", inplace=True)
    df5["地区编码"].replace("南京刘纪彬", "0403", inplace=True)
    df5["负责人"].replace("南京刘纪彬", "刘纪彬", inplace=True)
    df5["部门"].replace("南京刘纪彬", "04部", inplace=True)

    df5["地区"].replace("南京陈豪", "南京4", inplace=True)
    df5["地区编码"].replace("南京陈豪", "0404", inplace=True)
    df5["负责人"].replace("南京陈豪", "陈豪", inplace=True)
    df5["部门"].replace("南京陈豪", "04部", inplace=True)

  ###五部
    df5["地区"].replace("南通朱一亦", "南通1", inplace=True)
    df5["地区编码"].replace("南通朱一亦", "0501", inplace=True)
    df5["负责人"].replace("南通朱一亦", "李国旺 朱一亦", inplace=True)
    df5["部门"].replace("南通朱一亦", "05部", inplace=True)

    df5["地区"].replace("南通王峥骅", "南通2", inplace=True)
    df5["地区编码"].replace("南通王峥骅", "0502", inplace=True)
    df5["负责人"].replace("南通王峥骅", "李国旺 王铮骅", inplace=True)
    df5["部门"].replace("南通王峥骅", "05部", inplace=True)

    df5["地区"].replace("盐城", "盐城", inplace=True)
    df5["地区编码"].replace("盐城", "0503", inplace=True)
    df5["负责人"].replace("盐城", "岑潭泽 潘前进", inplace=True)
    df5["部门"].replace("盐城", "05部", inplace=True)

    df5["地区"].replace("连云港", "连云港", inplace=True)
    df5["地区编码"].replace("连云港", "0504", inplace=True)
    df5["负责人"].replace("连云港", "胡士艳 姜健", inplace=True)
    df5["部门"].replace("连云港", "05部", inplace=True)

  ###六部

    df5["地区"].replace("上海1", "上海1", inplace=True)
    df5["地区编码"].replace("上海1", "0601", inplace=True)
    df5["负责人"].replace("上海1", "汤俊", inplace=True)
    df5["部门"].replace("上海1", "06部", inplace=True)

    df5["地区"].replace("上海2", "上海2", inplace=True)
    df5["地区编码"].replace("上海2", "0602", inplace=True)
    df5["负责人"].replace("上海2", "邬幼波", inplace=True)
    df5["部门"].replace("上海2", "06部", inplace=True)

  ####七部

    df5["地区"].replace("苏州", "苏州", inplace=True)
    df5["地区编码"].replace("苏州", "0701", inplace=True)
    df5["负责人"].replace("苏州", "陈凯", inplace=True)
    df5["部门"].replace("苏州", "07部", inplace=True)

    df5["地区"].replace("苏州市郊", "苏郊", inplace=True)
    df5["地区编码"].replace("苏州市郊", "0702", inplace=True)
    df5["负责人"].replace("苏州市郊", "吕楠", inplace=True)
    df5["部门"].replace("苏州市郊", "07部", inplace=True)


   ####八部

    df5["地区"].replace("扬泰1", "扬泰1", inplace=True)
    df5["地区编码"].replace("扬泰1", "0801", inplace=True)
    df5["负责人"].replace("扬泰1", "姜海涛", inplace=True)
    df5["部门"].replace("扬泰1", "08部", inplace=True)

    df5["地区"].replace("扬泰2", "扬泰2", inplace=True)
    df5["地区编码"].replace("扬泰2", "0802", inplace=True)
    df5["负责人"].replace("扬泰2", "吕淳昱", inplace=True)
    df5["部门"].replace("扬泰2", "08部", inplace=True)

    df5["地区"].replace("扬泰3", "扬泰3", inplace=True)
    df5["地区编码"].replace("扬泰3", "0803", inplace=True)
    df5["负责人"].replace("扬泰3", "胡霞", inplace=True)
    df5["部门"].replace("扬泰3", "08部", inplace=True)


   ####九部

    df5["地区"].replace("徐州于博", "徐州", inplace=True)
    df5["地区编码"].replace("徐州于博", "0901", inplace=True)
    df5["负责人"].replace("徐州于博", "唐维洲 于博", inplace=True)
    df5["部门"].replace("徐州于博", "09部", inplace=True)

    df5["地区"].replace("徐州张浩", "徐州", inplace=True)
    df5["地区编码"].replace("徐州张浩", "0901", inplace=True)
    df5["负责人"].replace("徐州张浩", "唐维洲 于博", inplace=True)
    df5["部门"].replace("徐州张浩", "09部", inplace=True)

    df5["地区"].replace("宿迁", "宿迁", inplace=True)
    df5["地区编码"].replace("宿迁", "0902", inplace=True)
    df5["负责人"].replace("宿迁", "赵晨阳 王涛", inplace=True)
    df5["部门"].replace("宿迁", "09部", inplace=True)

    df5["地区"].replace("淮安", "淮安", inplace=True)
    df5["地区编码"].replace("淮安", "0903", inplace=True)
    df5["负责人"].replace("淮安", "赵晨阳 白虹", inplace=True)
    df5["部门"].replace("淮安", "09部", inplace=True)


    ####十部

    df5["地区"].replace("常州", "常州", inplace=True)
    df5["地区编码"].replace("常州", "1001", inplace=True)
    df5["负责人"].replace("常州", "吴羚", inplace=True)
    df5["部门"].replace("常州", "10部", inplace=True)

    df5["地区"].replace("镇江", "镇江", inplace=True)
    df5["地区编码"].replace("镇江", "1002", inplace=True)
    df5["负责人"].replace("镇江", "周丹", inplace=True)
    df5["部门"].replace("镇江", "10部", inplace=True)

    df5["地区"].replace("常镇", "常镇", inplace=True)
    df5["地区编码"].replace("常镇", "1003", inplace=True)
    df5["负责人"].replace("常镇", "于亚惠", inplace=True)
    df5["部门"].replace("常镇", "10部", inplace=True)

   ####十一部

    df5["地区"].replace("绍兴龚群波", "绍兴", inplace=True)
    df5["地区编码"].replace("绍兴龚群波", "1101", inplace=True)
    df5["负责人"].replace("绍兴龚群波", "龚群波", inplace=True)
    df5["部门"].replace("绍兴龚群波", "11部", inplace=True)

    df5["地区"].replace("衢州", "金衢", inplace=True)
    df5["地区编码"].replace("衢州", "1102", inplace=True)
    df5["负责人"].replace("衢州", "胡迪锋", inplace=True)
    df5["部门"].replace("衢州", "11部", inplace=True)

    df5["地区"].replace("金华", "金衢", inplace=True)
    df5["地区编码"].replace("金华", "1102", inplace=True)
    df5["负责人"].replace("金华", "胡迪锋", inplace=True)
    df5["部门"].replace("金华", "11部", inplace=True)

   ####十二部
    df5["地区"].replace("无锡裘涌", "无锡1", inplace=True)
    df5["地区编码"].replace("无锡裘涌", "1201", inplace=True)
    df5["负责人"].replace("无锡裘涌", "裘涌", inplace=True)
    df5["部门"].replace("无锡裘涌", "12部", inplace=True)

    df5["地区"].replace("无锡张立伟、赵飞", "无锡2", inplace=True)
    df5["地区编码"].replace("无锡张立伟、赵飞", "1202", inplace=True)
    df5["负责人"].replace("无锡张立伟、赵飞", "张立伟 赵飞", inplace=True)
    df5["部门"].replace("无锡张立伟、赵飞", "12部", inplace=True)

    ####调拨

    df5["地区"].replace("北京", "调拨", inplace=True)
    df5["地区编码"].replace("北京", "1301", inplace=True)
    df5["负责人"].replace("北京", "孙婷婷", inplace=True)
    df5["部门"].replace("北京", "调拨", inplace=True)

    df5["地区"].replace("生化分销", "调拨", inplace=True)
    df5["地区编码"].replace("生化分销", "1301", inplace=True)
    df5["负责人"].replace("生化分销", "孙婷婷", inplace=True)
    df5["部门"].replace("生化分销", "调拨", inplace=True)

    df5["地区"].replace("调拨", "调拨", inplace=True)
    df5["地区编码"].replace("调拨", "1301", inplace=True)
    df5["负责人"].replace("调拨", "孙婷婷", inplace=True)
    df5["部门"].replace("调拨", "调拨", inplace=True)

    df5["地区"].replace("维修部", "调拨", inplace=True)
    df5["地区编码"].replace("维修部", "1301", inplace=True)
    df5["负责人"].replace("维修部", "孙婷婷", inplace=True)
    df5["部门"].replace("维修部", "调拨", inplace=True)

    df6 = df5[df5["部门"] != "关联企业"]



    tkinter.messagebox.showinfo("提醒", "请选择英克销售发票明细（新试剂）源文件");
    # 加入开票明细

    df51 = pd.read_excel(tkinter.filedialog.askopenfilename());  # 630修改为手工读路径

    df52 = df51.fillna(0)
    df53 = df52.drop(df52.index[[0, 1]], axis=0);

    df53["仪器开票合计"] = df53["03原厂仪器"] + df53["04采购平台仪器"] + df53["0501国产辅助配置"] + df53["0502流水线辅助配置"]

    df53["非仪器开票合计"] = df53["010101免疫（代理）"] + df53["010102特定蛋白（代理）"] + df53["010103血球（代理）"] + df53["010104普通生化（代理）"] \
                     + df53["010105AU生化（代理）"] + df53["010106利德曼生化（代理）"] + df53["010107尿液（代理）"] + df53["010109微生物（代理）"] + \
                     df53["010110索灵（代理BC）"] \
                     + df53["010111免疫（AMH）"] + df53["010201血凝（代理）"] + df53["0103lmmucor"] + df53["0104索灵"] + df53[
                         "010501质控试剂（代理）"] + df53["010502伯乐其它试剂"] \
                     + df53["010601BNP试剂（代理）"] + df53["010701血气（代理）"] + df53["0108苏医（代理BC血球质控）"] + df53["020101干式生化"] + df3[
                         "020102普通生化"] \
                     + df53["020103血气"] + df53["020104特殊生化"] + df53["020201血球"] + df53["020202血凝"] + df53["020203尿液"] + df53[
                         "020204血库"] + df53["020206体液"] \
                     + df53["020301发光"] + df53["020302特定蛋白"] + df53["020303酶免类"] + df53["020304其它免疫"] + df53["020305厦门万泰"] + \
                     df53["0204微生物"] + df53["0205药字号"] \
                     + df53["0206分子诊断"] + df53["0207病理科"] + df53["0208采购平台其它"] + df53["0209质控"] + df53["06软件"] + df53["07配件"] \
                     + df53["08其它业务"] + df53["0901基因试剂（自产）"] + df53["0902基因试剂（其它厂家）"] + df53["1101强盛生化"] + df53[
                         "1201沃文特免疫"] + df53["1202沃文特其他"] \
                     + df53["99其它"]

    df54 = df53.drop(["Unnamed: 1", "Unnamed: 2", "Unnamed: 3"], axis=1)  # 删列

    df55 = df54.groupby(["Unnamed: 0"], as_index=False)["非仪器开票合计", "仪器开票合计"].sum();

    df55["地区"] = df55["Unnamed: 0"]
    df55["负责人"] = df55["Unnamed: 0"]
    df55["地区编码"] = df55["Unnamed: 0"]
    df55["部门"] = df55["Unnamed: 0"]

    df55["地区"].replace("温州葛瑞", "温州1", inplace=True)
    df55["地区编码"].replace("温州葛瑞", "0101", inplace=True)
    df55["负责人"].replace("温州葛瑞", "葛瑞", inplace=True)
    df55["部门"].replace("温州葛瑞", "01部", inplace=True)

    df55["地区"].replace("台州唐惠", "台州1", inplace=True)
    df55["地区编码"].replace("台州唐惠", "0103", inplace=True)
    df55["负责人"].replace("台州唐惠", "唐惠", inplace=True)
    df55["部门"].replace("台州唐惠", "01部", inplace=True)

    df55["地区"].replace("温州潘磊", "温州2", inplace=True)
    df55["地区编码"].replace("温州潘磊", "0102", inplace=True)
    df55["负责人"].replace("温州潘磊", "潘磊", inplace=True)
    df55["部门"].replace("温州潘磊", "01部", inplace=True)

    df55["地区"].replace("台州胡文魁", "台州2", inplace=True)
    df55["地区编码"].replace("台州胡文魁", "0104", inplace=True)
    df55["负责人"].replace("台州胡文魁", "胡文魁", inplace=True)
    df55["部门"].replace("台州胡文魁", "01部", inplace=True)

    df55["地区"].replace("丽水", "丽水", inplace=True)
    df55["地区编码"].replace("丽水", "0105", inplace=True)
    df55["负责人"].replace("丽水", "方汝泼", inplace=True)
    df55["部门"].replace("丽水", "01部", inplace=True)

    #####一部完毕
    df55["地区"].replace("宁波市区", "宁波", inplace=True)
    df55["地区编码"].replace("宁波市区", "0201", inplace=True)
    df55["负责人"].replace("宁波市区", "丁玲", inplace=True)
    df55["部门"].replace("宁波市区", "02部", inplace=True)

    df55["地区"].replace("舟山北仑", "舟山北仑", inplace=True)
    df55["地区编码"].replace("舟山北仑", "0202", inplace=True)
    df55["负责人"].replace("舟山北仑", "高大勇", inplace=True)
    df55["部门"].replace("舟山北仑", "02部", inplace=True)

    df55["地区"].replace("慈溪余姚镇海", "北三县", inplace=True)
    df55["地区编码"].replace("慈溪余姚镇海", "0203", inplace=True)
    df55["负责人"].replace("慈溪余姚镇海", "陆金耀", inplace=True)
    df55["部门"].replace("慈溪余姚镇海", "02部", inplace=True)

    df55["地区"].replace("奉化宁海象山", "南三县", inplace=True)
    df55["地区编码"].replace("奉化宁海象山", "0204", inplace=True)
    df55["负责人"].replace("奉化宁海象山", "吴燕江", inplace=True)
    df55["部门"].replace("奉化宁海象山", "02部", inplace=True)
    ####三部####

    df55["地区"].replace("杭州姜立民", "省级", inplace=True)
    df55["地区编码"].replace("杭州姜立民", "0301", inplace=True)
    df55["负责人"].replace("杭州姜立民", "姜立民", inplace=True)
    df55["部门"].replace("杭州姜立民", "03部", inplace=True)

    df55["地区"].replace("杭州石亚国", "省级", inplace=True)
    df55["地区编码"].replace("杭州石亚国", "0301", inplace=True)
    df55["负责人"].replace("杭州石亚国", "姜立民", inplace=True)
    df55["部门"].replace("杭州石亚国", "03部", inplace=True)

    df55["地区"].replace("杭州陈靓", "省级", inplace=True)
    df55["地区编码"].replace("杭州陈靓", "0301", inplace=True)
    df55["负责人"].replace("杭州陈靓", 
