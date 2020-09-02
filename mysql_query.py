#!/usr/bin/python
import pymysql
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook
import datetime

db_host = 'rm-bp1a237026czz5i900o.mysql.rds.aliyuncs.com'
db_username = 'chenhao'
db_password = 'urvgycQEVSx1RT2m'
db_name = 'pmtool'
report_status = []
report_manager = []
dead_line = []
# 打开数据库连接
db = pymysql.connect(db_host, db_username, db_password, db_name,
                     charset='utf8')
# 绩效文件名
fileSource = "每周日报.xlsx"


def get_days_daily(days_ago):
    now_time = datetime.datetime.now()

    delta_day = datetime.timedelta(days=-days_ago + 1)

    da_days = (now_time + delta_day).strftime('%Y-%m-%d %000:%000:%000')

    # 使用cursor()方法获取操作游标
    cursor = db.cursor()

    sql_str = 'SELECT pmtool.p_news.p_id ,p_name,pmtool.p_news.create_time,pmtool.p_news.p_news ,p_status,p_manager,p_design_time,p_codecomplete_time ,p_archive_time ,p_online_time FROM pmtool.pmlist,pmtool.p_news where pmtool.p_news.p_id = pmtool.pmlist.id  and pmtool.p_news.create_time >= ' + '\'' + da_days + '\''
    # sql_str = 'SELECT pmtool.p_news.p_id ,p_name,pmtool.p_news.create_time,pmtool.p_news.p_news ,p_status,p_manager FROM pmtool.pmlist,pmtool.p_news where pmtool.p_news.p_id = pmtool.pmlist.id  and pmtool.p_news.create_time and p_status <5 '

    print("sql_str=" + sql_str)

    # 使用execute方法执行SQL语句
    cursor.execute(sql_str)

    # 使用 fetchone() 方法获取一条数据
    data = cursor.fetchall()

    return data


def get_one_daily(date_start, data):
    report_map = {}
    for rows in data:
        try:
            d1 = rows[2].strftime('%Y-%m-%d')
            if date_start == d1:
                if report_map.get(rows[1]) is None:
                    report_map[rows[1]] = rows[3]
                else:
                    report_map[rows[1]] = report_map.get(rows[1]) + " &&&&&" + rows[3]

        except ValueError as e:
            print(e)
    return report_map


def get_project_name(data):
    report_name = []
    report_status.clear()
    for rows in data:
        name = rows[1]
        if name not in report_name:
            if rows[4] == 0:
                report_name.append(name)
                report_status.append('设计中')
                report_manager.append(rows[5])
                dead_line.append(str(rows[6]))
            if rows[4] == 1:
                report_name.append(name)
                report_status.append('开发中')
                report_manager.append(rows[5])
                dead_line.append(str(rows[7]))
            if rows[4] == 2:
                report_name.append(name)
                report_status.append('测试中')
                report_manager.append(rows[5])
                dead_line.append(str(rows[8]))
            if rows[4] == 3:
                report_name.append(name)
                report_status.append('申请归档')
                report_manager.append(rows[5])
                dead_line.append(str(rows[9]))

    return report_name


def get_project_status(data):
    report_name = []
    for rows in data:
        name = rows[1]
        if name not in report_name:
            report_name.append(name)

    return report_name


def get_daily_excel(days_no):
    data = get_days_daily(days_no);
    project_names = get_project_name(data)
    pro_dic = {'项目名称': project_names, '项目状态': report_status, '项目经理': report_manager, '阶段截止日期': dead_line}
    now_time = datetime.datetime.now()

    r = range(0, days_no)
    for i in r:
        delta_day = datetime.timedelta(days=(i + 1) - days_no)
        da_days = (now_time + delta_day).strftime('%Y-%m-%d')
        print(da_days)
        daily_map = get_one_daily(da_days, data)
        one_report = []
        for name in project_names:
            one_report.append(daily_map.get(name))
        pro_dic[da_days] = one_report
    return pro_dic


# 创建缓存区
writer = pd.ExcelWriter(fileSource)

book = load_workbook(fileSource)
writer.book = book
dic = get_daily_excel(60)

print(dic)
df1 = pd.DataFrame(dic)
# 写入sheet
df1.to_excel(writer, sheet_name='每周日报', index=False)

writer.save()
