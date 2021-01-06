# coding:utf-8
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook

# 绩效文件名
fileSource = "12月发布记录.xlsx"
# 创建缓存区
writer = pd.ExcelWriter(fileSource)
f = pd.ExcelFile(fileSource)
book = load_workbook(fileSource)
writer.book = book
# 创建临时姓名、未验证次数
dev_names = {}

for i in f.sheet_names:
    sheet = pd.read_excel(fileSource, sheet_name=i)
    df = pd.DataFrame(sheet)
    data = df.values
    for rows in data:
        if pd.isnull(rows[4]):
            dev_name = rows[3]
        else:
            dev_name = rows[4]
        if pd.isnull(rows[7]):
            tes_name = rows[6]
        else:
            tes_name = rows[7]
        print(rows[1])
        print(tes_name)
        if not pd.isnull(tes_name):
            tes_name = tes_name.strip()
        if not pd.isnull(dev_name):
            dev_name = dev_name.strip()
        if dev_names.get(dev_name) is None:
            if not pd.isnull(rows[5]):
                dev_names[dev_name] = 0
        if dev_names.get(dev_name) is not None:
            if not pd.isnull(rows[5]):
                dev_names[dev_name] = dev_names.get(dev_name) + 1
        if dev_names.get(tes_name) is None:
            if not pd.isnull(rows[8]) and pd.isnull(rows[9]):
                dev_names[tes_name] = 1
        if dev_names.get(tes_name) is not None:
            if not pd.isnull(rows[8]) and pd.isnull(rows[9]):
                dev_names[tes_name] = dev_names.get(tes_name) + 1

# del dev_names['nan']
# del tes_names['nan']
print("姓名：")
print(dev_names)
final_dev_names = []
final_dev_credits = []
for key in dev_names.keys():
    if not pd.isna(key):
        final_dev_names.append(key)
        final_dev_credits.append(dev_names.get(key))
dic = {'姓名': final_dev_names, '合计': final_dev_credits}
df1 = pd.DataFrame(dic)
df1 = df1.sort_values(['合计'], ascending=False)
# 写入sheet
df1.to_excel(writer, sheet_name='发布绩效', index=False)

writer.save()
