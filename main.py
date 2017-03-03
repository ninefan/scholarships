# -*- coding:utf-8 -*-

import xlrd
import xlwt
from xlutils.copy import copy

TOTAL_MONEY = 200000
SINGLE_MONEY = [2000, 500, 500, 500, 500]


class Row:
    def __init__(self, grade, score, data):
        self.grade = grade
        self.score = score
        self.data = data
        self.score_sum = 0

    def set_score_sum(self, score_sum):
        self.score_sum = score_sum

    # 比较函数，先比较年级小在前，后比较分数高在前
    def __lt__(self, other):
        if self.grade != other.grade:
            return self.grade < other.grade
        if self.score != other.score:
            return self.score > other.score
        return self.score_sum > other.score_sum


# 读入xlsx
workbook = xlrd.open_workbook(u'2016奖学金打分汇总表.xlsx')
# 写入xls
workbook_ = copy(workbook)
sheet_names = workbook.sheet_names()
print(sheet_names)

# 每张表中成绩对应的列数
score_col = [0, 5, 8, 11, 4, 4, 4, 4, 4, 4]
# 学生姓名和成绩的映射关系
score_map = {}


def get_score(_table, _table_id):
    _score = _table.cell(i, score_col[_table_id]).value
    try:
        _score = int(_score)
    except ValueError:
        _score = 0
    return _score


def get_row(_table, _i, _score, class_id_col=1):
    class_id = _table.cell(_i, class_id_col).value
    if len(class_id) == 0:
        grade = -1
    else:
        grade = class_id[1:2]
    data = []
    for _j in range(0, table.ncols):
        data.append(table.cell(_i, _j).value)
    return Row(grade, _score, data)


table_id = 1
table = workbook.sheet_by_index(table_id)

money_multiple = 0

for i in range(2, table.nrows):
    name = table.cell(i, 0).value
    if len(name) == 0:
        continue
    if score_map.get(name) is None:
        score_map[name] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        money = table.cell(i, 9).value
        try:
            money = int(money)
        except ValueError:
            money = 0
        money_multiple += money
        score_map[name][table_id] = money

# 分析后八张表的内容
sort_list_arr = []
for table_id in range(2, 10):
    table = workbook.sheet_by_index(table_id)
    table_ = workbook_.get_sheet(table_id)
    sort_list = []
    print(table.nrows)
    for i in range(2, table.nrows):
        name = table.cell(i, 0).value
        if len(name) > 0:
            score = get_score(table, table_id)
            if score_map.get(name) is not None:
                # 获取并设置分数
                score_map[name][table_id] = score
            if score_map.get(name) is None or score_map[name][1] == 0:
                # 获取年级并加入总排序表
                row = get_row(table, i, score)
                sort_list.append(row)
        for j in range(0, table.ncols):
            table_.write(i, j, None)

    # 排序并输出
    sort_list.sort()
    sort_list_arr.append(sort_list)
    for i in range(len(sort_list)):
        print(sort_list[i].data)
        for j in range(len(sort_list[i].data)):
            table_.write(i + 2, j, sort_list[i].data[j])

table_id = 1
table = workbook.sheet_by_index(table_id)
table_ = workbook_.get_sheet(table_id)

# 遍历综合优秀表
sort_list = []
for i in range(2, table.nrows):
    name = table.cell(i, 0).value
    if len(name) == 0:
        continue

    score = get_score(table, table_id)
    row = get_row(table, i, score, 3)
    score_sum = 0

    for j in range(2, 10):
        row.data[12 + j] = score_map[name][j]
        score_sum += score_map[name][j]

    row.score_sum = row.data[6] = score_sum
    sort_list.append(row)

# 排序并输出
sort_list.sort()

for i in range(len(sort_list)):
    for j in range(len(sort_list[i].data)):
        table_.write(i + 2, j, sort_list[i].data[j])


def create_sub_table(_name, _col):
    _table_ = workbook_.add_sheet(_name)
    _table_.write(0, 0, _name)
    for _i in range(0, table.ncols):
        _table_.write(1, _i, table.cell(1, _i).value)
    _row_num = 1
    for _row in sort_list:
        if _row.data[_col] == u'是':
            _row_num += 1
            for _i in range(len(_row.data)):
                _table_.write(_row_num, _i, _row.data[_i])


create_sub_table(u'贫困', 10)
create_sub_table(u'少数名族', 11)

table_ = workbook_.add_sheet(u'奖金统计')
table_.write(0, 0, u'奖金统计')
for i in range(0, 5):
    table_.write(0, i + 1, 10 - i)
table_.write(0, 6, u'总计奖金')

money_single = 0

for table_id in range(2, 10):
    sort_list = sort_list_arr[table_id - 2]
    table_.write(table_id - 1, 0, sheet_names[table_id])
    people_sum = [0, 0, 0, 0, 0]
    money_row = 0
    for row in sort_list:
        if 6 <= row.score <= 10:
            people_sum[10 - row.score] += 1
            money_row += SINGLE_MONEY[10 - row.score]
    for i in range(0, 5):
        table_.write(table_id - 1, i + 1, people_sum[i])
    table_.write(table_id - 1, 6, money_row)
    money_single += money_row

table_.write(11, 0, u'总奖金')
table_.write(11, 1, TOTAL_MONEY)
table_.write(12, 0, u'综合优秀')
table_.write(12, 1, money_multiple)
table_.write(13, 0, u'单项奖金')
table_.write(13, 1, money_single)
table_.write(14, 0, u'剩余')
table_.write(14, 1, TOTAL_MONEY - money_multiple - money_single)

workbook_.save('test.xls')
