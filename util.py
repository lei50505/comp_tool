#! /usr/bin/env python
# -*- coding: utf-8 -*-

'''doc'''

def singleton(cls):
    '''doc'''
    instances = {}
    def _singleton(*args, **kw):
        if cls not in instances:
            instances[cls] = cls(*args, **kw)
        return instances[cls]
    return _singleton

@singleton
class Data():
    '''doc'''
    def __init__(self):
        self.data = {}
    def set(self, key, val):
        '''doc'''
        self.data[key] = val
    def get(self, key):
        '''doc'''
        return self.data[key]

def to_float(val):
    '''doc'''
    try:
        return float(val)
    except:
        return None

def to_str(val):
    '''doc'''
    if val is None:
        return None
    try:
        val = str(val)
        if val.strip() == "":
            return None
        else:
            return val.strip()
    except:
        return None
