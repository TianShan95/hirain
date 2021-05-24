#!/usr/bin/python
# -*- coding: utf-8 -*-
from globalvar import Var  # 导入全局字典
from tkinter import *


class Win:

    def __init__(self):
        self.excelFile_suffix = Var.get_value('excelFile_suffix')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.excelFile_Path = Var.get_value('excelFile_Path')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.excelFile = Var.get_value('excelFile')  # 在字典中取出之前存入的excelFile_suffix(列表)
        self.Done_excelFile = Var.get_value('Done_excelFile')  # 在字典中取出之前存入的excelFile_suffix(列表)

    def Clear_list(self):
        listbox = Var.get_value('listbox')
        listbox.delete(0, END)

        self.excelFile_suffix.clear()  # 清空列表
        self.excelFile_Path.clear()
        self.excelFile.clear()
        self.Done_excelFile.clear()

    def Clear_info(self):
        text = Var.get_value('text')
        text.config(state=NORMAL)
        text.delete(0.0, END)
        text.config(state=DISABLED)
