# -*- coding:utf-8 -*-

import xlrd
import xlwt
from xlutils.copy import copy


class Row:
    def __init__(self, grade, score, data):
        self.grade = grade
        self.score = score
        self.data = data

    # 比较函数，先比较年级小在前，后比较分数高在前
    def __lt__(self, other):
        if self.grade == other.grade:
            return self.score > other.score
        return self.grade < other.grade

# 读入xlsx
workbook = xlrd.open_workbook(u'2016奖学金打分汇总表.xlsx')
# 写入xls
workbook_ = copy(workbook)

print(workbook.sheet_names())

# 每张表中成绩对应的列数
score_col = [0, 0, 8, 11, 4, 4, 4, 4, 4, 4]
# 学生姓名和成绩的映射关系
score_map = {}

# 分析后八张表的内容
for table_id in range(2, 10):
    table = workbook.sheet_by_index(table_id)
    sort_list = []
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

        # 获取年级并加入总排序表
        class_id = table.cell(i, 1).value
        if len(class_id) == 0:
            continue
        grade = class_id[1:2]
        data = []
        for j in range(0, table.ncols):
            data.append(table.cell(i, j).value)
        sort_list.append(Row(grade, score, data))

    # 排序并输出
    sort_list.sort()
    table_ = workbook_.get_sheet(table_id)
    for i in range(len(sort_list)):
        for j in range(len(sort_list[i].data)):
            table_.write(i + 2, j, sort_list[i].data[j])

table = workbook.sheet_by_index(1)
table_ = workbook_.get_sheet(1)

# 遍历综合优秀表
for i in range(2, table.nrows):
    name = table.cell(i, 0).value
    if len(name) == 0:
        continue
    # print(name)
    if score_map.get(name) is None:
        score_map[name] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    for j in range(2, 10):
        table_.write(i, 12 + j, score_map[name][j])

workbook_.save('test.xls')
