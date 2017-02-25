# -*- coding:utf-8 -*-

import xlrd
import xlwt
a=[[0]*5]*1000;
workbook = xlrd.open_workbook(u'2016奖学金打分汇总表.xlsx')
print(workbook.sheet_names())

table = workbook.sheet_by_name(u'学业优秀')
print(table.nrows, table.ncols)

# 这张表从第3行开始，29行结束
# 相当于for(i=2;i<29;i++)
for i in range(2, 28):
    # 第1列为姓名
    a[i][1]=table.cell(i, 0).value
    a[i][2]=table.cell(i, 8).value


#workbook1 = xlwt.open_workbook(u'2016奖学金打分汇总表.xlsx')
#table1 = workbook1.sheet_by_name(u'综合优秀')
#for i in range(2,29):
#    table.write(i,15,a[i][2])

