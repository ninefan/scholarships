import xlrd, xlwt

workbook = xlrd.open_workbook('2016奖学金打分汇总表.xlsx')
print(workbook.sheet_names())

table = workbook.sheet_by_name('学业优秀')
print(table.nrows, table.ncols)

# 这张表从第3行开始，29行结束
# 相当于for(i=2;i<29;i++)
for i in range(2, 29):
    # 第1列为姓名
    print(table.cell(i, 0).value)

