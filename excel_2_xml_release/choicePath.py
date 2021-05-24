#!/usr/bin/python3
# coding=utf-8
from tkinter import *
from tkinter import filedialog
import os  # 用于查找路径下的目标文件
from globalvar import Var  # 导入全局字典


class ChoicePath:
    def __init__(self):
        self.excelFile_suffix = Var.get_value('excelFile_suffix')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.excelFile_Path = Var.get_value('excelFile_Path')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.excelFile = Var.get_value('excelFile')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.Done_excelFile = Var.get_value('Done_excelFile')  # 在字典中取出之前存入的excelFile_suffix(列表)

    def file_name(self, file_dir):

        listbox = Var.get_value('listbox')  # 在字典中取出之前存入的listbox
        listbox_row = Var.get_value('listbox_row')

        for root, dirs, files in os.walk(file_dir):
            for file in files:
                if os.path.splitext(file)[1] == '.xlsx':  # 需要的是后缀为xlsx的表格
                    (path, filename) = os.path.split(file)  # 分离路径和文件名
                    filename = os.path.splitext(filename)[0]  # 分离出不带后缀的文件名(有个bug是把第一个点以前的字符算作是不带后缀的文件名)
                    self.excelFile_suffix.append(root + '/' + file)  # 存入文件路径+文件名
                    self.excelFile.append(filename)  # 存入不带后缀的文件名
                    self.excelFile_Path.append(root)  # 存入每个需要处理文件的路径
                    listbox.insert(END, str(len(self.excelFile_suffix)) + ' : ' + root + '/' + file + "   Added\n")
                    listbox.yview(MOVETO, 1.0)
                    listbox_row = listbox_row + 1
                    Var.set_value('listbox_row', listbox_row)  # 将State存入全局变量,以便在其他py文件调用配置
        return self.excelFile_suffix

    def choicePath(self):
        foldPath = filedialog.askdirectory(title=u"Please choice a folder")  # 所选择文件夹的路径

        self.file_name(foldPath)  # 匹配这个文件夹路径下以xlsx为后缀的表格

        Var.set_value('excelFile_suffix', self.excelFile_suffix)  # 将excelFile_suffix存入全局变量,以便在其他py文件调用配置
        Var.set_value('excelFile_Path', self.excelFile_Path)  # 将excelFile_Path存入全局变量,以便在其他py文件调用配置
        Var.set_value('excelFile', self.excelFile)  # 将excelFile存入全局变量,以便在其他py文件调用配置
        Var.set_value('Done_excelFile', self.Done_excelFile)  # 将Done_excelFile存入全局变量,以便在其他py文件调用配置
