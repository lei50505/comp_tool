#! /usr/bin/env python
# -*- coding: utf-8 -*-

'''doc'''
# pylint:disable=too-many-branches,too-many-statements,too-many-locals,broad-except

import time
import traceback

from book import Book
from util import Data

def main():
    '''doc'''

    in_book = Book("in.xlsx")

    print("正在读取in.xlsx")
    in_book.load()

    if not in_book.has_sheet("Sheet1"):
        print("in.xlsx中不存在Sheet1")
        in_book.close()
        return

    if not in_book.has_sheet("Sheet2"):
        print("in.xlsx中不存在Sheet2")
        in_book.close()
        return

    in_sheet1 = in_book.get_sheet("Sheet1")
    in_sheet2 = in_book.get_sheet("Sheet2")

    data = Data()

    print("获取Sheet1数字列序号")
    sheet1_num_col_index = in_sheet1.get_num_col_index()
    if data.get("ERR"):
        print(data.get("ERR_MSG"))
        in_book.close()
        return

    print("获取Sheet2数字列序号")
    sheet2_num_col_index = in_sheet2.get_num_col_index()
    if data.get("ERR"):
        print(data.get("ERR_MSG"))
        in_book.close()
        return

    out_book = Book("out.xlsx")
    out_book.create()

    out_sheet = out_book.get_active_sheet()

    sheet1_row_done = {}
    sheet2_row_done = {}

    sheet1_max_row = in_sheet1.get_max_row()
    sheet2_max_row = in_sheet2.get_max_row()

    for row_index in range(1, sheet1_max_row + 1):
        sheet1_row_done[row_index] = False

    for row_index in range(1, sheet2_max_row + 1):
        sheet2_row_done[row_index] = False


    print("获取Sheet1数字列数据")
    sheet1_num_col_data = in_sheet1.get_num_col_data()

    print("获取Sheet2数字列数据")
    sheet2_num_col_data = in_sheet2.get_num_col_data()

    print("获取Sheet1数字唯一行")
    sheet1_alone_val_rows = in_sheet1.get_alone_val_rows()

    print("获取Sheet2数字唯一行")
    sheet2_alone_val_rows = in_sheet2.get_alone_val_rows()

    sheet1_row_val_dict = sheet1_num_col_data.get("ROW_VAL_DICT")
    sheet2_row_val_dict = sheet2_num_col_data.get("ROW_VAL_DICT")
    sheet1_val_set = sheet1_num_col_data.get("VAL_SET")
    sheet2_val_set = sheet2_num_col_data.get("VAL_SET")
    sheet1_val_cnt_dict = sheet1_num_col_data.get("VAL_CNT_DICT")
    sheet2_val_cnt_dict = sheet2_num_col_data.get("VAL_CNT_DICT")

    print("正在处理数字唯一行")
    for sheet1_row_index in sheet1_alone_val_rows:
        for sheet2_row_index in sheet2_alone_val_rows:
            sheet1_val = sheet1_row_val_dict[sheet1_row_index]
            sheet2_val = sheet2_row_val_dict[sheet2_row_index]

            if sheet1_val == sheet2_val:
                sheet1_row_done[sheet1_row_index] = True
                sheet2_row_done[sheet2_row_index] = True

                out_sheet.copy_row_from_sheet(in_sheet1, sheet1_row_index, "BLUE")
                out_sheet.copy_row_from_sheet(in_sheet2, sheet2_row_index, "RED")

    print("正在处理数字个数相同行")
    cnt_equal_vals = []
    for sheet1_val in sheet1_val_set:
        sheet1_count = sheet1_val_cnt_dict[sheet1_val]
        if sheet1_count == 1:
            continue
        for sheet2_val in sheet2_val_set:
            sheet2_count = sheet2_val_cnt_dict[sheet2_val]
            if sheet2_count == 1:
                continue

            if sheet1_val != sheet2_val:
                continue

            if sheet1_count != sheet2_count:
                continue

            cnt_equal_vals.append(sheet1_val)

    for val in cnt_equal_vals:

        for sheet1_row_index in in_sheet1.get_rows_by_val(val):

            sheet1_row_done[sheet1_row_index] = True

            out_sheet.copy_row_from_sheet(in_sheet1, sheet1_row_index, "BLUE")


        for sheet2_row_index in in_sheet2.get_rows_by_val(val):

            sheet2_row_done[sheet2_row_index] = True

            out_sheet.copy_row_from_sheet(in_sheet2, sheet2_row_index, "RED")


    print("正在处理没有匹配的行")
    row_data_list = []
    for sheet1_row_index in range(1, sheet1_max_row + 1):

        if sheet1_row_done[sheet1_row_index]:

            continue

        row_data = {}
        cell_val = in_sheet1.get_cell(sheet1_row_index, sheet1_num_col_index).get_float_val()
        if cell_val is None:
            cell_val = 0
        if cell_val < 0:
            cell_val = 0 - cell_val
        row_data["key"] = cell_val
        row_data["data"] = sheet1_row_index
        row_data_list.append(row_data)

    row_data_list = sorted(row_data_list, key=lambda item: item["key"])

    row_index_list = []
    for row_item in row_data_list:
        row_index_list.append(row_item["data"])


    for sheet1_row_index in row_index_list:
        out_sheet.copy_row_from_sheet(in_sheet1, sheet1_row_index, "BLUE")


    row_index_list = []
    row_data_list = []

    for sheet2_row_index in range(1, sheet2_max_row + 1):

        if sheet2_row_done[sheet2_row_index]:

            continue

        row_data = {}
        cell_val = in_sheet2.get_cell(sheet2_row_index, sheet2_num_col_index).get_float_val()
        if cell_val is None:
            cell_val = 0
        if cell_val < 0:
            cell_val = 0 - cell_val
        row_data["key"] = cell_val
        row_data["data"] = sheet2_row_index
        row_data_list.append(row_data)
    row_data_list = sorted(row_data_list, key=lambda item: item["key"])
    for row_item in row_data_list:
        row_index_list.append(row_item["data"])

    for sheet2_row_index in row_index_list:

        out_sheet.copy_row_from_sheet(in_sheet2, sheet2_row_index, "RED")

    out_book.save()

    in_book.close()

    out_book.close()
    print("处理成功")


if __name__ == "__main__":
    try:
        main()
        time.sleep(2)
    except Exception:
        print(traceback.format_exc())
        time.sleep(10)
