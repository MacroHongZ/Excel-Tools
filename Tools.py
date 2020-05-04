#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2020-03-10 13:42:46
# @Author  : mx (wanghongzhun@outlook.com)
# @Link    : ${link}
# @Version : $Id$

import tkinter as tk
from tkinter.filedialog import askopenfilenames,askopenfilename
from tkinter import messagebox
import pandas as pd
import xlrd

def get_ent():
	global col
	col = ent.get()
	ent.delete(0,"end")

def merge_excel(pathes):
	sheet = []
	for filepath in pathes:
		# sheet.append(pd.read_excel(filepath,header=2,converters = {u'学号':str}))
		sheet.append(pd.read_excel(filepath))
	he = pd.concat(sheet)
	he.to_excel("合并.xlsx",index=False)
	messagebox.showinfo('提示','合并完成')

def merge_sheet(path):
	sheet = []
	wb = xlrd.open_workbook(path)
	sheets = wb.sheet_names()
	for i in range(len(sheets)):
		sheet.append(pd.read_excel(path,sheet_name=i))
	he = pd.concat(sheet)
	he.to_excel("合并.xlsx",index=False)
	messagebox.showinfo('提示','合并完成')

def split_sheet(path):
		global col
		date = pd.read_excel(path,heder=0)
		col = eval(col)-1
		split_list = list(set(date.iloc[:,col]))
		writer = pd.ExcelWriter("拆分.xlsx")
		for j in split_list:
			df = date[date.iloc[:,col]==j]
			df.to_excel(writer,sheet_name=j,index=False)
		writer.save()
		messagebox.showinfo('提示','拆分完成')

def selectpath1():
	pathes = askopenfilenames()
	merge_excel(pathes)

def selectpath2():		
	path = askopenfilename()
	merge_sheet(path)

def selectpath3():
	if col == "wu":
		messagebox.showinfo('提示','请先在右侧输入要拆分的列')
	else:	
		path = askopenfilename()
		split_sheet(path)


col = "wu"
window = tk.Tk()
window.title('Excel Tools')
# 获取屏幕 宽、高
ws = window.winfo_screenwidth()
hs = window.winfo_screenheight()
# 计算 x, y 位置
w = 500
h = 300
x = int((ws/2)-(w/2))
y = int((hs/2)-(h/2))
window.geometry('{}x{}+{}+{}'.format(w,h,x,y))

fra = tk.Frame(window)
fra.place(x=50,y=50)
fra1 = tk.Frame(window)
fra1.place(x=150,y=50)

ent = tk.Entry(fra1,width=30)
button_ent = tk.Button(fra1,text='确定',width=4,height=1,command=get_ent)
labe2 = tk.Label(fra1,text='请输入要拆分的列')
labe2.pack()
ent.pack()
button_ent.pack()
# text = tk.Text(fra, height=20, width=50)
# text.pack(side='right')

button_1 = tk.Button(fra,text='合并多簿',width=10,height=3,bg='Chocolate',fg='GhostWhite',command=selectpath1)
button_1.pack()
button_2 = tk.Button(fra,text='合并多表',width=10,height=3,bg='Chocolate',fg='GhostWhite',command=selectpath2)
button_2.pack()
button_3 = tk.Button(fra,text='拆分到多表',width=10,height=3,bg='Chocolate',fg='GhostWhite',command=selectpath3)
button_3.pack()



# 第6步，主窗口循环显示
window.mainloop()







