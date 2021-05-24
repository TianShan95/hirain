#!/usr/bin/python3
# coding=utf-8
from tkinter import *
from tkinter import ttk
import time
import write_xml
from globalvar import Var  # 导入全局字典


class Convert:

    def __init__(self):
        self.excelFile_suffix = Var.get_value('excelFile_suffix')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.excelFile_Path = Var.get_value('excelFile_Path')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.excelFile = Var.get_value('excelFile')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.Done_excelFile = Var.get_value('Done_excelFile')  # 在字典中取出之前存入的excelFile_suffix(列表)

    def convert(self):

        win = Var.get_value('win')  # 在字典中取出之前存入的win
        convert = Var.get_value('convert')  # 在字典中取出之前存入的convert
        convert.place_forget()  # 隐藏转换按钮
        State_01 = Var.get_value('State_01')  # 在字典中取出之前存入的State（第几个/共几个）
        listbox = Var.get_value('listbox')
        listbox_row = Var.get_value('listbox_row')

        # 定义一种进度条格式
        bar_01 = ttk.Style()
        bar_01.theme_use('clam')
        curWidth = win.winfo_width()  # 获取窗口的宽度
        bar_01.configure("green.Horizontal.TProgressbar", foreground='red', background='green')
        mpb = ttk.Progressbar(win, style="green.Horizontal.TProgressbar", orient="horizontal", length=curWidth,
                              mode="determinate", maximum=100, value=0)
        mpb.place(relx=0, rely=0.9)

        print(len(self.excelFile_suffix))
        file_num = len(self.excelFile_suffix)  # 获取一共有多少需要转换的文件

        State_01.config(text='0/' + str(file_num), fg='green')
        State_01.place(relx=0.5, rely=0.95)
        win.update()  # 刷新页面
        time.sleep(0.1)

        for index, value in enumerate(self.excelFile_suffix):
            Num = index
            print(self.excelFile_suffix)
            print(self.excelFile_Path)
            print(self.excelFile)
            print(Num)
            print(value)
            State_01.config(text=str(Num) + '/' + str(file_num))
            win.update()  # 刷新页面
            time.sleep(0.1)
            result = write_xml.Excel2Xml(value, Num, self.excelFile_Path,
                                         self.excelFile, self.Done_excelFile, self.excelFile_suffix).read_excel()
            mpb["value"] = mpb["value"] + 100 / file_num
            listbox.delete(listbox_row - file_num + Num)
            if result == 'success':
                listbox.insert(listbox_row - file_num + Num, str(Num + 1) + ' : ' + value + "   Successfully\n")
                listbox.itemconfig(listbox_row - file_num + Num, fg='green')
            if result == 'fail':
                listbox.insert(listbox_row - file_num + Num, str(Num + 1) + ' : ' + value + "   Fail\n")
                listbox.itemconfig(listbox_row - file_num + Num, fg='blue')
            listbox.yview(MOVETO, 1.0)
            win.update()  # 刷新页面
            time.sleep(0.1)

        self.excelFile_suffix.clear()  # 清空列表
        self.excelFile_Path.clear()
        self.excelFile.clear()
        self.Done_excelFile.clear()
        State_01.config(text="DONE!", fg='green')
        win.update()  # 刷新页面
        time.sleep(0.5)
        State_01.place_forget()  # 隐藏提示信息
        mpb.place_forget()  # 隐藏屏幕的进度条
        convert.place(relx=0.45, rely=0.9)  # 打包select_02按钮,设置按钮的一些显示属性
        time.sleep(0.1)
