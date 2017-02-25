# -*- coding:utf-8 -*-

import xlrd
import xlwt
from xlutils.copy import copy

a = [[0] * 5] * 1000

workbook = xlrd.open_workbook(u'2016奖学金打分汇总表.xlsx')
workbook_ = copy(workbook)

print(workbook.sheet_names())

# 每张表中成绩对应的列数
score_col = [0, 0, 8, 11, 4, 4, 4, 4, 4, 4]
# 学生姓名和成绩的映射关系
score_map = {}

# 分析后八张表的内容
for table_id in range(2, 10):
    table = workbook.sheet_by_index(table_id)
    for i in range(2, table.nrows):
        name = table.cell(i, 0).value
        if len(name) == 0:
            continue
        if score_map.get(name) is None:
            score_map[name] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        # 获取并设置分数
        score = table.cell(i, score_col[table_id]).value
        try:
            score = int(score)
        except ValueError:
            score = 0
        score_map[name][table_id] = score

table = workbook.sheet_by_name(u'综合优秀')
table_ = workbook_.get_sheet(1)

# 遍历综合优秀表
for i in range(2, table.nrows):
    name = table.cell(i, 0).value
    print(name)
    if score_map.get(name) is None:
        score_map[name] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    for j in range(2, 10):
        table_.write(i, 12 + j, score_map[name][j])

workbook_.save('test.xls')
