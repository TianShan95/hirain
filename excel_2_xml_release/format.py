#!/usr/bin/python
# -*- coding: utf-8 -*-
from globalvar import Var  # 导入全局字典


class Format:

    def __init__(self):
        self.format = format

    def Matrix(self):
        self.format = Var.set_value('Format', 'Matrix')
        print(Var.get_value('Format'))

    def Event(self):
        self.format = Var.set_value('Format', 'BDCM event')
        print(Var.get_value('Format'))

    def Cycle(self):
        self.format = Var.set_value('Format', 'BDCM cycle')
        print(Var.get_value('Format'))
