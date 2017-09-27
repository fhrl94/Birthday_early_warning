import configparser
import datetime
import glob

import os

import sys
import zipfile

from sqlalchemy import text, and_, distinct

from emailtool import send, sendMultimedia
import time
import platform
import pymssql
from TimerTask import timer
from warnstone import EmployeeInfo, stoneobject, Relation, WeekMapping, MonthMapping


conf = configparser.ConfigParser()
if platform.system() == 'Windows':
    conf.read("warning.conf",encoding="utf-8-sig")
else:
    conf.read("warning.conf")
# 数据删除
def remove(stone):
    # 清空数据库
    stone.query(EmployeeInfo).delete()
    stone.query(Relation).delete()
    stone.query(WeekMapping).delete()
    stone.query(MonthMapping).delete()
    stone.commit()


# 数据初始化
# 1、从金蝶中获取职员表（职员信息（姓名、工号、出生日期）、职务、职位ID）
# 2、职位表（上下级关系）
def create_table(stone):
    #  将金蝶数据转存到sqlite 数据库中
    #  转存（在职人员） 职员表 姓名、工号、出生日期、职位唯一ID（涉及到sqlite支不支持UID），需要处理 职务的获取（存在兼职的情况）
    conn = pymssql.connect(conf.get('server', 'ip'), conf.get('server', 'user'), conf.get('server', 'password'),
                           database=conf.get('server', 'database'))
    cur = conn.cursor()
    sql = """
    select he.Name,he.Code,he.Birthday,ope.PositionID,oj.Name,ou.Name,ope.IsPrimary from HM_Employees as he 
    join ORG_Position_Employee as ope on he.EM_ID=ope.EmID 
    join ORG_Position as op on ope.PositionID=op.ID
    join ORG_Job as oj on op.JobID=oj.ID
    join ORG_Unit as ou on op.UnitID=ou.ID
    where he.Status =1 and op.IsDelete=0 and ou.StatusID=1"""
    cur.execute(sql)
    empcols = ['name', 'code', 'birthDate', 'positionID', 'job', 'departmentname', 'IsPrimary']
    for one in cur.fetchall():
        empinfo = EmployeeInfo()
        for count, col in enumerate(empcols):
            if col == 'positionID':
                setattr(empinfo, col, str(one[count]))
            else:
                setattr(empinfo, col, one[count])
        stone.add(empinfo)
    stone.commit()
    #  转存（存在的岗位） 职位表 当前职位ID，父级职位ID
    sql = """
    select ID,ParentID from ORG_Position"""
    cur.execute(sql)
    recols = ['positionID', 'parentID', ]
    print('_______')
    for one in cur.fetchall():
        print(one)
        relation = Relation()
        for count, col in enumerate(recols):
            setattr(relation, col, str(one[count]))
        stone.add(relation)
    print('_______')
    stone.commit()


# 数据转存
# 将指定日期期间的人员放入 转存表 的表头，通过迭代获取相应的上级
#  兼职人员处理
def unloading(stone, today, afterday, number, table):
    """

    :param stone: 数据库连接
    :param today:  当前日期
    :param afterday:  预警日期与当前日期相差天数
    :param number:  预警的周期天数
    :param table:  存储表
    :return:
    """
    #  日期处理迭代（可用函数实现），将时间区域的人转存到 月转存表 或 周转存表
    today = today + datetime.timedelta(days=afterday)
    for num in range(number):
        result = stone.query(EmployeeInfo).filter(
            and_(text("strftime('%m%d',DATE (birthDate,'1 day'))=strftime('%m%d',date(:date,:value)) "))
            , EmployeeInfo.IsPrimary == True).params(value='{num} day'.format(num=num + 1), date=today).all()
        for one in result:
            # print(one)
            tab = table()
            # 处理后缀数值
            try:
                int(one.name[len(one.name)-1])
                tab.name = one.name[:len(one.name)-1]
            except ValueError:
                tab.name = one.name
            tab.code = one.code
            tab.birthDate = one.birthDate
            tab.departmentname=one.departmentname
            tab.positionID = one.positionID
            tab.job = one.job
            tab.date = today + datetime.timedelta(days=num)
            tab.count = 0
            tab.director = None
            tab.director1 = None
            tab.manager = None
            tab.manager1 = None
            tab.majordomo = None
            tab.principal = None
            tab.general_manager = None
            stone.add(tab)
    stone.commit()
    #  上级获取迭代（可用函数实现）
    result = stone.query(table).all()
    for one in result:
        position = one.positionID
        director = []
        manager = []
        majordomo = []
        principal = []
        parent = stone.query(Relation).filter(Relation.positionID == position).one_or_none()
        while parent:
            position = parent.parentID
            parent = stone.query(EmployeeInfo).filter(EmployeeInfo.positionID == position).one_or_none()
            if parent is None:
                print('空')
                break
            if parent.job == '主管':
                director.append(parent.name)
            elif parent.job == '经理':
                manager.append(parent.name)
            elif parent.job == '总监':
                majordomo.append(parent.name)
            elif parent.job == '副总':
                principal.append(parent.name)
            parent = stone.query(Relation).filter(Relation.positionID == position).one_or_none()
        if director:
            one.director1 = director.pop()
        if director:
            one.director = director.pop()
        if manager:
            one.manager1 = manager.pop()
        if manager:
            one.manager = manager.pop()
        if majordomo:
            one.majordomo = majordomo.pop()
        if principal and principal[-1] != conf.get(section='special', option='name'):
            one.general_manager = principal.pop()
        if principal:
            one.principal = principal.pop()
    stone.commit()


# 偏函数
# 邮件发送
def email_send(stone, table,):
    """

    :param stone: 数据库连接
    :param table: 指定数据库
    :return: 无返回值
    """
    # TODO 获取发件人的邮箱、发件内容处理（无人员处理），生成正文内容body
    cols = ['director', 'director1', 'manager', 'manager1', 'majordomo', 'principal']
    starthtml = r"""<html>
    <body>"""
    # 生日表格
    shengri = """<table width="580"  border="1" align="center">
        <caption>
        生日名单
        </caption>
        <tbody>
            <tr>
                <th width="20%" scope="col">工号</th>
                <th width="20%" scope="col">部门名称</th>
                <th width="20%" scope="col">姓名</th>
                <th width="20%" scope="col">生日日期</th>
            </tr>"""
    endhtml = '  </tbody> </table>'
    # 删除上一次生成的所有文件
    dirlist = sys.path[0] + os.sep + 'temp' + os.sep
    for delfile in os.listdir(dirlist):
        os.remove(sys.path[0] + os.sep + 'temp' + os.sep + delfile)
    for col in cols:
        result = stone.query(distinct(getattr(table, col))).filter(getattr(table, col) != None).all()
        # print(result)
        for one in result:
            body = ''
            # print(one)
            # print(stone.query(table).filter(getattr(table,col)==one[0]).all())
            if col == 'principal':
                gather = stone.query(table).filter(
                    and_(getattr(table, col) == one[0], getattr(table, 'job') != '员工')).all()
            else:
                gather = stone.query(table).filter(getattr(table, col) == one[0]).all()
            for emp in gather:
                # print(emp)
                body = body + r'<tr>      <td width="20%" align="center">{code}</td>' \
                              r'<td width="20%" align="center">{departmentname}</td>' \
                              r'<td width="20%" align="center">{name}</td>' \
                              r'<td width="20%" align="center">{date}</td></tr>'.format(
                    code=emp.code, departmentname=emp.departmentname, name=emp.name, date=emp.date)
                emp.count +=1
                # print(body)
            # print(starthtml + shengri + body + endhtml)
            # TODO 邮件发送
            # send(smtp_server=smtp_server, smtp_port=smtp_port, from_addr=from_addr, from_addr_str=from_addr_str,
            #      password=password, to_address=to_address, header=header, body=starthtml + shengri + body + endhtml, )
            file = open('temp'+os.sep+one[0]+'.html', 'a')
            file.write(starthtml + shengri + body + endhtml)
            file.close()
    stone.commit()
    result=stone.query(table).filter(table.count == 0)
    for one in result:
        body = body + r'<tr>      <td width="20%" align="center">{code}</td>' \
                      r'<td width="20%" align="center">{departmentname}</td>' \
                      r'<td width="20%" align="center">{name}</td>' \
                      r'<td width="20%" align="center">{date}</td></tr>'.format(code=one.code,
                                                                                departmentname=one.departmentname,
                                                                                name=one.name, date=one.date)
    file = open('temp' + os.sep + '{today}未发送人员'.format(today=datetime.date.today()) + '.html', 'a')
    file.write(starthtml + shengri + body + endhtml)
    file.close()
    if os.path.exists(sys.path[0] + os.sep + r'发送情况.zip'):
        os.remove(sys.path[0] + os.sep + r'发送情况.zip')
    f = zipfile.ZipFile(sys.path[0] + os.sep + r'发送情况.zip', 'w', zipfile.ZIP_DEFLATED)
    files = glob.glob(sys.path[0] + os.sep + 'temp' + os.sep +'*')
    for file in files:
        f.write(file, os.path.basename(file))
    f.close()
    sendMultimedia(smtp_server=conf.get(section='email', option='smtp_server'), smtp_port=conf.get(
                   section='email', option='smtp_port'), from_addr=conf.get(section='email', option='from_addr')
                   , from_addr_str=conf.get(section='email', option='from_addr_str'), password=conf.get(
                   section='email', option='password'), to_address=conf.get(section='email', option='error_email')
                   , header='{today} 生日预警情况'.format(today=datetime.date.today()), body='邮件发送详情见附件'
                   , file=sys.path[0] + os.sep + u'发送情况.zip')


# 主程序
def main():
    #  初始化数据库连接
    stone = stoneobject()
    targettimestr = input('请输入定点时间，例如8:00')
    if targettimestr == '':
        targettimestr = conf.get(section='time', option='now')
    targettime = datetime.time(int(targettimestr.split(':')[0]), int(targettimestr.split(':')[1]))
    while True:
        if datetime.datetime.now().hour == targettime.hour:
            date = datetime.date.today() + datetime.timedelta(days=3)
            # TODO 判断时间，如果时间为月末最后一天，调用unloading 存到 月转存表
            if (date + datetime.timedelta(days=1)).month - date.month != 0:
                # 初始化数据库中的4个基本表
                remove(stone)
                # 数据提取
                create_table(stone)
                i=28
                while True:
                    if (date + datetime.timedelta(days=i+1)).month != (date + datetime.timedelta(days=1)).month:
                        break
                    i += 1
                unloading(stone, today=date, afterday=1, number=i, table=MonthMapping)
                email_send(stone, table=MonthMapping)
                print(i)
            #  判断时间，如果时间为周五，调用unloading 存到 周转存表
            # 不需要每周一次
            # if date.isoweekday() == 5:
            #     unloading(stone, today=date, afterday=3, number=7, table=WeekMapping)
            #     email_send(stone, table=WeekMapping, smtp_server='', smtp_port='', from_addr='', from_addr_str='',
            #                password='', to_address='', header='')
            # TODO 数据清理函数
            time.sleep(timer(targettime))
        else:
            print(timer(targettime))
            time.sleep(timer(targettime))
        pass
    pass


if __name__ == '__main__':
    main()
