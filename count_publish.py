# coding:utf-8
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook

# 绩效文件名
fileSource = "9月发布记录.xlsx"
# 创建缓存区
writer = pd.ExcelWriter(fileSource)
f = pd.ExcelFile(fileSource)
book = load_workbook(fileSource)
writer.book = book
# 创建临时姓名、未验证次数
publish_names = {}
for i in f.sheet_names:
    sheet = pd.read_excel(fileSource, sheet_name=i)
    df = pd.DataFrame(sheet)
    data = df.values
    for rows in data:
        publish_name = rows[0]
        if not pd.isnull(publish_name):
            if publish_names.get(publish_name) is None:
                publish_names[publish_name] = 0
            if publish_names.get(publish_name) is not None:
                publish_names[publish_name] = publish_names[publish_name] + 1
print(publish_names.__str__())
