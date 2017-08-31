#! /usr/bin/env python
# -*- coding: utf-8 -*-

'''doc'''

import time
from functools import wraps

def singleton(cls):
    '''doc'''
    instances = {}
    def _singleton(*args, **kw):
        if cls not in instances:
            instances[cls] = cls(*args, **kw)
        return instances[cls]
    return _singleton

def fn_time(func):
    '''doc'''
    @wraps(func)
    def function_time(*args, **kw):
        '''doc'''
        start_time = time.time()
        result = func(*args, **kw)
        end_time = time.time()
        print("[%s]:[%s seconds]" \
            % (func.__name__, str(end_time - start_time)))
        return result
    return function_time

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
        return self.data.get(key)

def to_float(val):
    '''doc'''
    # pylint:disable=broad-except
    try:
        return float(val)
    except Exception:
        return None

def to_str(val):
    '''doc'''
    # pylint:disable=broad-except
    if val is None:
        return None
    try:
        val = str(val)
        if val.strip() == "":
            return None
        return val.strip()
    except Exception:
        return None
