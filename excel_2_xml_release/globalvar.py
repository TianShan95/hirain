#!/usr/bin/python
# -*- coding: utf-8 -*-
import threading


class Dic:
    _instance_lock = threading.Lock()

    def __init__(self):
        self._global_dict = {}

    def __new__(cls, *args, **kwargs):
        if not hasattr(cls, '_instance'):
            with cls._instance_lock:  # 加锁
                cls._instance = super(Dic, cls).__new__(cls)
        return cls._instance

    def set_value(self, name, value):
        self._global_dict[name] = value

    def get_value(self, name):
        try:
            return self._global_dict[name]
        except KeyError:
            return None


Var = Dic()
