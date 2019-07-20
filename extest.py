#! /usr/bin/ python
# -*- coding: utf-8 -*-
#pyexcelの操作テスト

import pyexcel as ex

if  __name__ == '__main__':
    book = ex.PyExcel("test.xlsx")
    book.Help()
    
    #シート操作テスト
    sheet = book.GetSheet()#Error
    
    #シートの作成（操作シートの決定）・消去
    book.SetSheet("Sheet")
    book.RemoveSheet("Sheet")
    
    #シートの作成
    name = "TestSheet"
    book.SetSheet(name)
    sheet = book.GetSheet()
    
    
    #シートへの書き込み
    book.InsertSheet(3,3,"aiueo",1)
    book.InsertSheet(3,3,"aiueo2",1)
    book.InsertSheet(1,1,"test")
    book.InsertSheet(1,2,"test2")
    sheet.cell(row=2,column=2).value="direct"#
    print book.GetSheetValue(1,1)
    print book.GetSheetValue(1,2)
    print book.GetSheetValue(2,2)
    print book.GetSheetValue(3,3)
