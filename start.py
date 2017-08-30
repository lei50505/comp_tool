#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys

import traceback

from src.Book import *
import time

def do():
    inBook=None
    outBook=None
    try:
        print("请等待...")

        inBook = loadBook("in.xlsx")
        
        if not inBook.hasSheet("Sheet1"):
            raise Exception("in.xlsx中不存在Sheet1")
        if not inBook.hasSheet("Sheet2"):
            raise Exception("in.xlsx中不存在Sheet2")

        inSheet1 = inBook.sheet("Sheet1")

        print("初始化Sheet1数字列序号")
        inSheet1.initNumColIndex()


        
        inSheet2 = inBook.sheet("Sheet2")


        print("初始化Sheet2数字列序号")
        inSheet2.initNumColIndex()

        
        outBook = createBook()
        
        outSheet = outBook.active()


        sheet1RowDone = {}
        sheet2RowDone = {}

        sheet1MaxRow = inSheet1.getMaxRow()
        sheet2MaxRow = inSheet2.getMaxRow()
        sheet1MaxCol = inSheet1.getMaxCol()
        sheet2MaxCol = inSheet2.getMaxCol()
        
        for rowIndex in range(1, sheet1MaxRow + 1):
            sheet1RowDone[rowIndex] = False

        for rowIndex in range(1, sheet2MaxRow + 1):
            sheet2RowDone[rowIndex] = False


        print("初始化Sheet1数字列数据")
        inSheet1.initNumColDict()

        print("初始化Sheet2数字列数据")
        inSheet2.initNumColDict()


        print("初始化Sheet1唯一数字列")
        inSheet1.initDiffNumRows()


        print("初始化Sheet2唯一数字列")
        inSheet2.initDiffNumRows()


        print("正在处理唯一不同的树值")
        for sheet1RowIndex in inSheet1.diffNumRows:

            if sheet1RowDone[sheet1RowIndex] == True:
                continue

            for sheet2RowIndex in inSheet2.diffNumRows:
                if sheet2RowDone[sheet2RowIndex] == True:
                    continue
                
                

                sheet1Num = inSheet1.numRowDict[sheet1RowIndex]
                sheet2Num = inSheet2.numRowDict[sheet2RowIndex]
                

                if sheet1Num == sheet2Num:
                    sheet1RowDone[sheet1RowIndex] = True
                    sheet2RowDone[sheet2RowIndex] = True

                    outSheet.copyRowFromSheet(inSheet1,sheet1RowIndex,"blue")
                    outSheet.copyRowFromSheet(inSheet2,sheet2RowIndex,"red")
                                        
                   

        vals = []

        

        for sheet1Val in inSheet1.numValSet:
            sheet1Count = inSheet1.numValDict[sheet1Val]
            if sheet1Count ==1:
                continue
            for sheet2Val in inSheet2.numValSet:
                sheet2Count = inSheet2.numValDict[sheet2Val]
                if sheet2Count ==1:
                    continue

                if sheet1Val != sheet2Val:
                    continue

                if sheet1Count != sheet2Count:
                    continue

                vals.append(sheet1Val)


        print("正在处理相同的树值")
        for val in vals:


            for sheet1RowIndex in inSheet1.getRowListByVal(val):
                if sheet1RowDone[sheet1RowIndex] == True:
                    
                    continue

                sheet1RowDone[sheet1RowIndex] = True

                outSheet.copyRowFromSheet(inSheet1,sheet1RowIndex,"blue")
                       

            for sheet2RowIndex in inSheet2.getRowListByVal(val):
                if sheet2RowDone[sheet2RowIndex] == True:
                    
                    continue
               
                sheet2RowDone[sheet2RowIndex] = True

                outSheet.copyRowFromSheet(inSheet2,sheet2RowIndex,"red")                        
                    
                    
        
        print("正在处理没有匹配的项目")


        rowDataList = []
        for sheet1RowIndex in range(1,sheet1MaxRow+1):

            if sheet1RowDone[sheet1RowIndex] == True:
                
                continue

            rowData = {}
            cellVal = inSheet1.cell(sheet1RowIndex, inSheet1.numColIndex).getFloatVal()
            if cellVal is None:
                cellVal = 0
            if cellVal < 0:
                cellVal = 0 - cellVal
            rowData["key"] = cellVal
            rowData["data"] = sheet1RowIndex
            rowDataList.append(rowData)

        rowDataList = sorted(rowDataList, key=lambda item:item["key"])
        
        rowIndexList = []
        for rowItem in rowDataList:
            rowIndexList.append(rowItem["data"])


        for sheet1RowIndex in rowIndexList:
        

            sheet1RowDone[sheet1RowIndex] = True
            outSheet.copyRowFromSheet(inSheet1,sheet1RowIndex,"blue")
                    

        rowIndexList = []
        rowDataList = []

        for sheet2RowIndex in range(1,sheet2MaxRow+1):

            if sheet2RowDone[sheet2RowIndex] == True:
                
                continue
            rowData = {}
            cellVal = inSheet2.cell(sheet2RowIndex, inSheet2.numColIndex).getFloatVal()
            if cellVal is None:
                cellVal = 0
            if cellVal < 0:
                cellVal = 0 - cellVal
            rowData["key"] = cellVal
            rowData["data"] = sheet2RowIndex
            rowDataList.append(rowData)
        rowDataList = sorted(rowDataList, key=lambda item:item["key"])
        for rowItem in rowDataList:
            rowIndexList.append(rowItem["data"])

        for sheet2RowIndex in rowIndexList:
                      
            sheet2RowDone[sheet2RowIndex] = True
            outSheet.copyRowFromSheet(inSheet2,sheet2RowIndex,"red")  
        
        
        outBook.save("out.xlsx")

        time.sleep(2)
    except Exception:
        print(traceback.format_exc())
        time.sleep(200)
        
    finally:
        if inBook is not None:
            inBook.close()
        if outBook is not None:
            outBook.close()

            
        
        

    
        
    
if __name__=="__main__":
    do()

