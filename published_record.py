# coding:utf-8
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook

# 绩效文件名
fileSource = "7月发布记录.xlsx"
# 创建缓存区
writer = pd.ExcelWriter(fileSource)
f = pd.ExcelFile(fileSource)
book = load_workbook(fileSource)
writer.book = book
# 创建临时姓名、未验证次数
dev_names = {}
tes_names = {}

for i in f.sheet_names:
    sheet = pd.read_excel(fileSource, sheet_name=i)
    df = pd.DataFrame(sheet)
    data = df.values
    for rows in data:
        dev_name=rows[3]
        tes_name=rows[6]
        if dev_names.get(dev_name) is None:
            if pd.isnull(rows[5]):
                dev_names[dev_name]= 1
        if dev_names.get(dev_name) is not None:
            if pd.isnull(rows[5]):
                dev_names[dev_name]= dev_names.get(dev_name)+1
        if tes_names.get(tes_name) is None:
            if pd.isnull(rows[8]):
                tes_names[tes_name]= 1
        if tes_names.get(tes_name) is not None:
            if pd.isnull(rows[8]):
                tes_names[tes_name]= tes_names.get(tes_name)+1

#del dev_names['nan']
#del tes_names['nan']
print("开发：")
print(dev_names)
print("测试：")
print(tes_names)

final_dev_names = []
final_dev_credits = []
final_tes_names = []
final_tes_credits = []
for key in dev_names.keys():
    if not pd.isna(key):
        final_dev_names.append(key)
        final_dev_credits.append(dev_names.get(key))
for key in tes_names.keys():
    if not pd.isna(key):
        final_tes_names.append(key)
        final_tes_credits.append(tes_names.get(key))


dic = {'开发姓名': final_dev_names, '合计': final_dev_credits}
dic2 = {'测试姓名': final_tes_names, '合计': final_tes_credits}

df1 = pd.DataFrame(dic)
df1 = df1.sort_values(['合计'], ascending=False)

df2 = pd.DataFrame(dic2)
df2 = df2.sort_values(['合计'], ascending=False)

# 写入sheet
df1.to_excel(writer, sheet_name='开发合计', index=False)
df2.to_excel(writer, sheet_name='测试合计', index=False)
writer.save()