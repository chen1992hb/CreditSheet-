# coding:utf-8
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook

# 绩效文件名
fileSource = "技术部项目奖金分配202012.xlsx"
# 创建缓存区
writer = pd.ExcelWriter(fileSource)
book = load_workbook(fileSource)
writer.book = book


def count_publish(sheet_name):
    # 读取sheet
    sheet = pd.read_excel(fileSource, sheet_name=sheet_name)
    if sheet is None:
        return
    df = pd.DataFrame(sheet)
    date = df.values
    # 创建临时姓名、绩效list
    names = []
    credit_list = []
    credit_map = {}
    for rows in date:
        if not pd.isnull(rows[0]) and not pd.isnull(rows[1]):
            credit_map[rows[0]] = int(int(rows[1])*0.25)

    return credit_map


# excel地址
def count(sheet_name, result_sheet_name):
    # 读取sheet
    sheet = pd.read_excel(fileSource, sheet_name=sheet_name)
    if sheet is None:
        return
    df = pd.DataFrame(sheet)
    date = df.values
    # 创建临时姓名、绩效list
    names = []
    credit_list = []
    credit_map = {}
    # 遍历和赋值
    for rows in date:
        if rows[0] == '姓名':
            for name in rows:
                if not pd.isnull(name):
                    names.append(name)

        if str(rows[0]).strip() == '核准':
            for credit in rows:
                if not pd.isnull(credit):
                    credit_list.append(credit)
    print(names)
    print(credit_list)
    # 建立洗牌后和合计的姓名数据列表
    true_name = []
    true_credits = []
    print(names.__len__())
    print(credit_list.__len__())
    for index in range(len(names)):
        name = names[index]
        credit = credit_list[index]
        if type(credit) == int and type(name) != int and not pd.isnull(name):
            if type(credit_map.get(name)) == int and type(name) != int:
                credit_map[name] = credit_map.get(name) + credit
            else:
                credit_map[name] = credit
    print(result_sheet_name)
    print(credit_map)
    for key in credit_map.keys():
        true_name.append(key)
        true_credits.append(credit_map[key])

    # 格式化数据
    dic = {'姓名': true_name, '合计': true_credits}
    df1 = pd.DataFrame(dic)
    # 写入sheet
    df1.to_excel(writer, sheet_name=result_sheet_name, index=False)
    writer.save()
    return credit_map


project_map = count(sheet_name='项目', result_sheet_name='项目统计')
fu_map = count(sheet_name='FU+', result_sheet_name='Fu+统计')
data_map = count(sheet_name='数据', result_sheet_name='数据统计')
service_map = count(sheet_name='对接', result_sheet_name='对接统计')
publish_map=count_publish(sheet_name='发布绩效')


for key in project_map.keys():
    if fu_map.get(key) is not None:
        sumCredit = project_map.get(key) + fu_map.get(key)
        project_map[key] = sumCredit
        del fu_map[key]

    if data_map.get(key) is not None:
        sumCredit = project_map.get(key) + data_map.get(key)
        project_map[key] = sumCredit
        del data_map[key]

    if service_map.get(key) is not None:
        sumCredit = project_map.get(key) + service_map.get(key)
        project_map[key] = sumCredit
        del service_map[key]


    if publish_map.get(key) is not None:
        sumCredit = project_map.get(key) + publish_map.get(key)
        project_map[key] = sumCredit
        del publish_map[key]

print(fu_map)
print(data_map)
print(service_map)
for key in fu_map.keys():
    if project_map.get(key) is None:
        project_map[key] = fu_map.get(key)
    else:
        project_map[key] = fu_map.get(key) + project_map[key]

for key in data_map.keys():
    if project_map.get(key) is None:
        project_map[key] = data_map.get(key)
    else:
        project_map[key] = data_map.get(key) + project_map[key]

for key in service_map.keys():
    if project_map.get(key) is None:
        project_map[key] = service_map.get(key)
    else:
        project_map[key] = service_map.get(key) + project_map[key]

for key in publish_map.keys():
    if project_map.get(key) is None:
        project_map[key] = publish_map.get(key)
    else:
        project_map[key] = publish_map.get(key) + project_map[key]

print("合计绩效")
print(project_map)

final_names = []
final_credits = []

for key in project_map.keys():
    final_names.append(key)
    final_credits.append(project_map.get(key))

dic = {'姓名': final_names, '合计': final_credits}

df1 = pd.DataFrame(dic)
df1 = df1.sort_values(['合计'], ascending=False)
# 写入sheet
df1.to_excel(writer, sheet_name='合计绩效', index=False)
writer.save()

# pandas处理NaN - 参考链接：https://www.jianshu.com/p/41039996d867
# pandas追加sheet到已有sheet的excel文件会覆盖的问题 - 参考链接：https://blog.csdn.net/qq_44315987/article/details/104100281
# pandas排序 - 参考链接：https://www.jianshu.com/p/b2f414c50d0c
