#!/usr/bin/python3
# coding=utf-8
from tkinter import *
from tkinter import filedialog  # 选择文件的对话框
import os  # 用于查找路径下的目标文件
from globalvar import Var  # 导入全局字典


class ChoiceFiles:

    def __init__(self):
        self.excelFile_suffix = Var.get_value('excelFile_suffix')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.excelFile_Path = Var.get_value('excelFile_Path')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.excelFile = Var.get_value('excelFile')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.Done_excelFile = Var.get_value('Done_excelFile')  # 在字典中取出之前存入的excelFile_suffix(列表)

    def choiceFiles(self):  # 点击choiceFiles的函数

        listbox = Var.get_value('listbox')  # 在字典中取出之前存入的listbox
        text = Var.get_value('text')  # 在字典中取出之前存入的text
        times = Var.get_value('times')  # 在字典中取出之前存入的listbox
        listbox_row = Var.get_value('listbox_row')

        if_add_file = 0  # if_add_file表示是否在list里添加了新的文件名
        repeat_fileName = 0  # 选择的文件名在文件名list是否有重复
        add_fileName = filedialog.askopenfilename(title=u"Please choice your excel file")  # 获取鼠标手动选择的路径+文件名
        print(add_fileName)
        (file, ext) = os.path.splitext(add_fileName)  # 把路径+文件名分离出后缀名,并将后缀字符串传给ext
        (path, filename) = os.path.split(add_fileName)  # 把路径+文件名分离路径和文件名,路径传给path,文件名传给filename
        filename = os.path.splitext(filename)[0]  # 把filename分离出不带后缀的文件名
        for fileName in self.excelFile_suffix:  # 遍历列表里是否出现了本次选择的文件,避免重复添加
            if fileName == add_fileName:
                repeat_fileName = 1  # 如果该文件之前添加过,则标志位置1
        if add_fileName != '' and repeat_fileName == 0 and ext == '.xlsx':  # 如果选择了文件名，并且选择的和之前的文件不同，则把路径加如列表
            self.excelFile_suffix.append(add_fileName)  # 存入的是路径+文件名完整信息
            self.excelFile_Path.append(path)  # 存入此次选择的文件路径
            self.excelFile.append(filename)  # 存入不带后缀名的文件名
            if_add_file = 1  # 成功添加了文件

        if if_add_file == 1:  # 如果成功添加了文件
            if_add_file = 0  # 清标志位
            print(self.excelFile_suffix)

            # text.config(state=NORMAL)
            # text.insert('insert', '>>> valid File\n', 'tag')  # 申明使用tag中的设置 绿色
            # text.yview(MOVETO, 1.0)
            # text.config(state=DISABLED)

            listbox.insert(END, str(len(self.excelFile_suffix)) + ' : ' + add_fileName + "   Added\n")  # 插入选项
            listbox.yview(MOVETO, 1.0)
            listbox_row = listbox_row + 1
            Var.set_value('listbox_row', listbox_row)  # 将State存入全局变量,以便在其他py文件调用配置

        else:
            text.config(state=NORMAL)
            text.insert('insert', '>>>' + add_fileName + ' Not a valid file\n', 'tag_02')
            text.yview(MOVETO, 1.0)
            text.config(state=DISABLED)
            print(self.excelFile_suffix)

        Var.set_value('excelFile_suffix', self.excelFile_suffix)  # 将excelFile_suffix存入全局变量,以便在其他py文件调用配置
        Var.set_value('excelFile_Path', self.excelFile_Path)  # 将excelFile_Path存入全局变量,以便在其他py文件调用配置
        Var.set_value('excelFile', self.excelFile)  # 将excelFile存入全局变量,以便在其他py文件调用配置
        Var.set_value('Done_excelFile', self.Done_excelFile)  # 将Done_excelFile存入全局变量,以便在其他py文件调用配置
        Var.set_value('times', times)
    
        print(if_add_file)  # 打印需要处理的文件名
