# coding:utf-8
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook
import datetime

# 绩效文件名
fileSource = "12345.xlsx"
# 创建缓存区
writer = pd.ExcelWriter(fileSource)
book = load_workbook(fileSource)
writer.book = book

sheet = pd.read_excel(fileSource)
df = pd.DataFrame(sheet)
d2 = datetime.datetime.strptime('2020-04-26 00:00:00', '%Y-%m-%d %H:%M:%S')
d3 = datetime.datetime.strptime('2020-04-27 00:00:00', '%Y-%m-%d %H:%M:%S')
d4 = datetime.datetime.strptime('2020-04-28 00:00:00', '%Y-%m-%d %H:%M:%S')

date = df.values
project_id = []


def get_one_daily(date_start, date_end):
    report_map = {}
    for rows in date:
        try:
            d1 = datetime.datetime.strptime(rows[2], '%Y-%m-%d %H:%M:%S')
            if date_start < d1 < date_end:
                if report_map.get(rows[1]) is None:
                    report_map[rows[1]] = rows[3]
                else:
                    report_map[rows[1]] = report_map.get(rows[1]) + " & " + rows[3]

        except ValueError as e:
            print(e)
    return report_map


day_one=get_one_daily(d2, d3)

day_two=get_one_daily(d3, d4)

final_names = []
one_report = []
two_report = []

for key in day_one.keys():
    final_names.append(key)
    one_report.append(day_one.get(key))
    two_report.append(day_two.get(key))

dic = {'项目': final_names, '2020-04-26': one_report,'2020-04-27':two_report}

df1 = pd.DataFrame(dic)
# 写入sheet
df1.to_excel(writer, sheet_name='日报', index=False)
writer.save()