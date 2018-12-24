#!/user/bin/env python
#-*- coding:utf-8 -*-
import xlrd
import xlwt

if __name__=="__main__":
    #excel文件全路径
    excelpath = "F:\\大会房间统计表.xls"
    #用于读取excel文件
    tableopen = xlrd.open_workbook(excelpath)
    #获取excel工作薄数
    count = len(tableopen.sheets())
    print("工作薄数为%s"%count)
    #获取表数据的行、列数
    table = tableopen.sheet_by_name(sheet_name='Sheet1')
    h = table.nrows
    l = table.ncols
    print("表数据的行数为%s,列数为%s"%(h,l))
    #循环读取数据
    for i in range(0,h):
        rowValues = table.row_values(i)#按行读取数据
        #输出读取的数据
        for data in rowValues:
            print(data)
        print("===============")