#!/usr/bin/python
from typing import List

import pymysql
import pandas as pd
import numpy as np
import openpyxl
import calendar
from openpyxl import load_workbook
import datetime

db_host = 'rm-bp1a237026czz5i900o.mysql.rds.aliyuncs.com'
db_username = 'chenhao'
db_password = 'urvgycQEVSx1RT2m'
db_name = 'pmtool'

# 打开数据库连接
db = pymysql.connect(db_host, db_username, db_password, db_name,
                     charset='utf8')
# 项目列表
fileSource = "9月第2周人员占用.xlsx"

project_list = []
requirement_list = []
requirement_titles = []
stand_by = []
design_list = []
designer = ["", ""]
develop_list = []
developer = []
test_list = []
tester = []
checking_list = []
ending_list = []
pause_list = []
online_list = []


class Project(object):
    def __init__(self, p_name, p_status, p_rank, p_design_time, p_codecomplete_time, p_archive_time, p_online_time,
                 p_manager: str, p_productmanager: str, p_designer: str, p_technical: str, p_test: str, p_coder: str):
        self.p_name = p_name
        self.p_status = p_status
        self.p_rank = p_rank
        self.p_design_time = p_design_time
        self.p_codecomplete_time = p_codecomplete_time
        self.p_archive_time = p_archive_time
        self.p_online_time = p_online_time
        self.p_productmanager = p_productmanager
        self.p_designer = p_designer
        self.p_technical = p_technical
        self.p_test = p_test
        self.p_coder = p_coder
        self.p_manager = p_manager


class Requirement(object):
    def __init__(self, title, rank, status, manager):
        self.title = title
        self.rank = rank
        self.status = status
        self.manager = manager


def get_level(self: Project):
    if self.p_rank == -1:
        return '【C级】'
    if self.p_rank == 0:
        return '【B级】'
    if self.p_rank == 1:
        return '【A级】'
    if self.p_rank == 2:
        return '【S级】'
    if self.p_rank == 3:
        return '【SS级】'
    if self.p_rank is None:
        return ' '


def get_status(self: Project):
    if self.p_status == 0:
        return '设计'
    if self.p_status == 1:
        return '开发'
    if self.p_status == 2:
        return '测试'
    if self.p_status == 3:
        return '归档'
    if self.p_status == 4:
        return '需求验收'


def get_member(self: Project, status: int):
    if status == 0:
        return self.p_designer.replace(',', "\r\n").replace('-技术部', "").replace('-产品部', "")
    if status == 1:
        return self.p_coder.replace(',', "\r\n").replace('-技术部', "").replace('-产品部', "")
    if status == 2:
        return self.p_test.replace(',', "\r\n").replace('-技术部', "").replace('-产品部', "")
    if status == 3:
        return ''
    if status == 4:
        return ''


def get_next_sunday():
    today = datetime.date.today()

    one_day = datetime.timedelta(days=1)

    m1 = calendar.SUNDAY
    i = 0;

    while i < 2:
        today += one_day
        if today.weekday() == m1:
            i += 1
    # print(str(today))
    return today

#china 0.053
#American 0.021
def get_next_month():
    today = datetime.datetime.now()
    first_day = datetime.datetime(today.year, today.month)
    return first_day


def get_week_day(date):
    week_day_dict = {
        0: '周一',
        1: '周二',
        2: '周三',
        3: '周四',
        4: '周五',
        5: '周六',
        6: '周日',
    }
    day = date.weekday()

    return week_day_dict[day]


def get_deadline(self):
    if self.p_status == 0:
        if self.p_design_time is None:
            return "[未填写]"
        else:
            if self.p_design_time <= get_next_sunday():
                return str(self.p_design_time)[5:] + "[" + get_week_day(self.p_design_time) + "]"
            else:
                return str(self.p_design_time)[5:]
    if self.p_status == 1:
        if self.p_codecomplete_time is None:
            return "[未填写]"
        else:
            if self.p_codecomplete_time <= get_next_sunday():
                return str(self.p_codecomplete_time)[5:] + "[" + get_week_day(self.p_codecomplete_time) + "]"
            else:
                return str(self.p_codecomplete_time)[5:]
    if self.p_status == 2:
        if self.p_archive_time is None:
            return "[未填写]"
        else:
            if self.p_archive_time <= get_next_sunday():
                return str(self.p_archive_time)[5:] + "[" + get_week_day(self.p_archive_time) + "]"
            else:
                return str(self.p_archive_time)[5:]
    if self.p_status == 3:
        if self.p_online_time is None:
            return "[未填写]"
        else:
            if self.p_online_time <= get_next_sunday():
                return str(self.p_online_time)[5:] + "[" + get_week_day(self.p_online_time) + "]"
            else:
                return str(self.p_online_time)[5:]


def get_projects(query: str):
    projects: List[Project] = []

    sql = 'SELECT p_name, p_status, p_rank, p_design_time, p_codecomplete_time, p_archive_time, p_online_time, p_manager, p_productmanager, p_designer, p_technical, p_test, p_coder FROM pmtool.pmlist where ' + query

    # 使用cursor()方法获取操作游标
    cursor = db.cursor()
    # 使用execute方法执行SQL语句
    cursor.execute(sql)
    # 使用 fetchone() 方法获取一条数据
    project_data = cursor.fetchall()
    cursor.close()
    for rows in project_data:
        p = Project(rows[0], rows[1], rows[2], rows[3], rows[4], rows[5], rows[6], rows[7], rows[8], rows[9], rows[10],
                    rows[11], rows[12])
        projects.append(p)

    return projects


def add_to_max(to_add_list: list, max_size):
    for one_data in to_add_list:
        while len(one_data) < max_size:
            one_data.append("")


def get_requirements(requirements_list):
    sql_requirement = 'SELECT title,rank,status,manager FROM pmtool.demand where status <2'

    cursor = db.cursor()
    # 使用execute方法执行SQL语句
    cursor.execute(sql_requirement)
    # 使用 fetchone() 方法获取一条数据
    requirements_data = cursor.fetchall()
    cursor.close()
    for rows in requirements_data:
        r = Requirement(rows[0], rows[1], rows[2], rows[3])
        requirements_list.append(r)


before_online = ' p_status <5 '

pause_query = ' p_status = 7'

#date_str = '\'%Y-%m\''
date_str = '\'%Y-%m\''

#online_query = 'p_real_online_time > date_format(now(),' + date_str + ') and p_status = 5 and p_online_time > date_format(now(),' + date_str + ')'

online_query = 'p_real_online_time < \'2021-04-01\' and p_status = 5 and p_real_online_time > \'2021-03-01\' '


project_list = get_projects(before_online)

pause_project = get_projects(pause_query)

online_project = get_projects(online_query)

online_manage_list = []

online_product_list = []

online_ui_list = []

online_develop_list = []

online_design_list = []

online_test_list =[]


get_requirements(requirement_list)

for requirement in requirement_list:
    if requirement.status == 0:
        requirement_titles.append(requirement.title)
    if requirement.status == 1:
        stand_by.append(requirement.title + '|' + requirement.manager)

for project in project_list:
    if project.p_status == 0:
        design_list.append(
            str(get_level(project))+ project.p_name + "\r\n" + project.p_manager.replace('-技术部', "").replace(
                '-产品部', "") + " " + str(get_deadline(project)))
        developer.append(get_member(project, 1))
    if project.p_status == 1:
        develop_list.append(
            str(get_level(project))+ project.p_name + "\r\n" + project.p_manager.replace('-技术部', "").replace(
                '-产品部', "") + " " + str(get_deadline(project)))
        tester.append(get_member(project, 2))
    if project.p_status == 2:
        test_list.append(
            str(get_level(project)) + project.p_name + "\r\n" + project.p_manager.replace('-技术部', "").replace(
                '-产品部', "") + " " + str(get_deadline(project)))
    if project.p_status == 3:
        checking_list.append(
            str(get_level(project))+ project.p_name + "\r\n" + project.p_manager.replace('-技术部', "").replace(
                '-产品部', "") + " " + str(get_deadline(project)))

for project in pause_project:
    pause_list.append(
        str(get_level(project))  + project.p_name + "\r\n" + project.p_manager.replace('-技术部', "").replace(
            '-产品部', ""))

for project in online_project:
    online_list.append(
        str(get_level(project)) + project.p_name + "\r\n" + project.p_manager.replace('-技术部', "").replace(
            '-产品部', ""))
    online_manage_list.append(project.p_manager.replace('-技术部', "").replace(
            '-产品部', ""))
    online_design_list.append(project.p_technical.replace('-技术部', "").replace(
            '-产品部', ""))
    online_develop_list.append(project.p_coder.replace('-技术部', "").replace(
            '-产品部', ""))
    online_product_list.append(project.p_productmanager.replace('-技术部', "").replace(
            '-产品部', ""))
    online_ui_list.append(project.p_designer.replace('-技术部', "").replace(
            '-产品部', ""))
    online_test_list.append(project.p_test.replace('-技术部', "").replace(
            '-产品部', ""))

requirement_titles_len = len(requirement_titles)
stand_by_len = len(stand_by)
design_list_len = len(design_list)
develop_list_len = len(develop_list)
test_list_len = len(test_list)
checking_list_len = len(checking_list)
pause_list_len = len(pause_list)
online_list_len = len(online_list)
online_manage_len = len(online_manage_list)
online_design_len = len(online_design_list)
online_develop_len = len(online_develop_list)
online_product_len = len(online_product_list)
online_ui_len = len(online_ui_list)

size_list = [requirement_titles_len, stand_by_len, len(designer), len(design_list), len(developer), develop_list_len,
             len(tester), len(test_list), checking_list_len, pause_list_len, online_list_len, online_manage_len,
             online_design_len, online_develop_len, online_product_len, online_ui_len,len(online_test_list)]
a = np.array(size_list)
max_len = a.max()

data = [requirement_titles, stand_by, designer, design_list, developer, develop_list,
        tester, test_list, checking_list, pause_list, online_list, online_manage_list, online_design_list,
        online_develop_list, online_product_list, online_ui_list,online_test_list]

add_to_max(data, max_len)

pro_dic = {'需求池:' + str(requirement_titles_len): requirement_titles, '待立项:' + str(stand_by_len): stand_by,
           '人员安排': designer, '设计中:' + str(design_list_len): design_list, "开发": developer,
           "开发中:" + str(develop_list_len): develop_list, "测试": tester, "测试中:" + str(test_list_len): test_list,
           '需求验收:' + str(checking_list_len): checking_list, "暂停：" + str(pause_list_len): pause_list,
           "本月上线项目：" + str(online_list_len): online_list, '项目经理': online_manage_list, '产品经理': online_product_list,
           '编码': online_develop_list, '技术方案': online_design_list, 'UI': online_ui_list,'test':online_test_list}

# 创建缓存区
writer = pd.ExcelWriter(fileSource)
book = load_workbook(fileSource)
writer.book = book

df1 = pd.DataFrame(pro_dic)
# 写入sheet
df1.to_excel(writer, sheet_name='8月第1周', index=False)
writer.save()

client_member = ['陈浩', '陈晨', '王庚', '王腾', '袁梦', '秦辉', '胡瑞', '张丰', '胡有明', '向健伟', '乔自强', '刘杰', '聂良', '朱光哲', '熊小辉', '童科',
                 '秦振磊']

server_member = ['郝龙潘', '陈辉', '甘琼', '郑文尧', '卢庆', '黄文聪', '喻磊', '马哲涛', '李梦琪', '黄威', '汪国兵', '陈龙', '郭欣怡', '艾青松', '孙剑波',
                 '余中伟', '郑敏', '黄杰']

project_member = ['贺攀', '靳坚', '马炬', '徐莎', '董迈克', '余中伟', '常如']

product_member = ['吴清子', '王学佳', '殷培培', '汪洁', '邹先铎', '朱超', '邓先宇', '董治伟', '陈瑾萱', '张岩', '成幸']

test_member = ['范琴', '王貂', '柳畅宇', '熊彬', '贺文颖', '熊应宏', '甘栋', '万苗']

project_member = ['马哲涛', '常如', '秦辉', '熊彬', '余中伟', '朱超', '邓先宇', '陈浩','张丰']


def save_member_sheet(member_list: list, sheet_name: str):
    dic = {}
    for member in member_list:
        member_projects = [member]
        for one_project in project_list:
            if one_project.p_coder.__contains__(member) and one_project.p_status == 1:
                member_projects.append(
                    str(get_level(
                        one_project)) + ")" + one_project.p_name + '\r\n[开发]-' + one_project.p_manager.replace(
                        '-技术部', "").replace(
                        '-产品部', "") + " " + str(get_deadline(one_project)))
                member_projects.append('')
            elif one_project.p_test.__contains__(member) and one_project.p_status == 2:
                member_projects.append(
                    str(get_level(
                        one_project)) + ")" + one_project.p_name + "\r\n[测试]-" + one_project.p_manager.replace(
                        '-技术部', "").replace(
                        '-产品部', "") + " " + str(get_deadline(one_project)))
                member_projects.append('')
            elif one_project.p_designer.__contains__(member) and one_project.p_status == 0:
                member_projects.append(
                    str(get_level(
                        one_project)) + ")" + one_project.p_name + '\r\n[设计]-' + one_project.p_manager.replace(
                        '-技术部', "").replace(
                        '-产品部', "") + " " + str(get_deadline(one_project)))
                member_projects.append('')
            elif one_project.p_productmanager.__contains__(member) and one_project.p_status == 0:
                member_projects.append(
                    str(get_level(
                        one_project)) + ")" + one_project.p_name + "\r\n[设计]-" + one_project.p_manager.replace(
                        '-技术部', "").replace(
                        '-产品部', "") + " " + str(get_deadline(one_project)))
                member_projects.append('')
            elif one_project.p_technical.__contains__(member) and one_project.p_status == 0:
                member_projects.append(
                    str(get_level(
                        one_project)) + ")" + one_project.p_name + "\r\n[方案]-" + one_project.p_manager.replace(
                        '-技术部', "").replace(
                        '-产品部', "") + " " + str(get_deadline(one_project)))
                member_projects.append('')
            elif \
                    one_project.p_manager.__contains__(member):
                member_projects.append(
                    str(get_level(
                        one_project)) + ")" + one_project.p_name + str(get_deadline(one_project)))
        dic[member] = member_projects
    print(dic)
    df1 = pd.DataFrame.from_dict(dic, orient='index')
    # 写入sheet
    df1.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.save()


# save_member_sheet(client_member, '客户端')
# save_member_sheet(server_member, '服务端')
# save_member_sheet(project_member, '项目部')
# save_member_sheet(product_member, '产品部')
# save_member_sheet(test_member, '测试部')
save_member_sheet(project_member, '项目部')
