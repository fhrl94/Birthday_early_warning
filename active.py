
import datetime

from sqlalchemy import text

from emailtool import send
import time
import pymssql
from TimerTask import timer
from warnstone import EmployeeInfo, stoneobject, Relation, WeekMapping, MonthMapping


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
    conn = pymssql.connect("地址", "用户名", "密码", database="账套号")
    cur = conn.cursor()
    sql = """
    select he.Name,he.Code,he.Birthday,ope.PositionID,oj.Name from HM_Employees as he 
    join ORG_Position_Employee as ope on he.EM_ID=ope.EmID 
    join ORG_Position as op on ope.PositionID=op.ID
    join ORG_Job as oj on op.JobID=oj.ID
    where Status =1 and ope.IsPrimary=1"""
    cur.execute(sql)
    empcols = ['name', 'code', 'birthDate', 'positionID', 'job']
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
# TODO 兼职人员处理
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
            text("strftime('%m%d',DATE (birthDate,'1 day'))=strftime('%m%d',date(:date,:value)) ")).params(
            value='{num} day'.format(num=num + 1), date=today).all()
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
        if principal:
            one.general_manager = principal.pop()
        if principal:
            one.principal = principal.pop()
    stone.commit()
    pass


# 偏函数
# 邮件发送
def email_send(smtp_server, smtp_port, from_addr, from_addr_str, password, to_address, header,):
    """

    :param smtp_server: SMTP地址
    :param smtp_port: SMTP端口
    :param from_addr: 发件人邮箱
    :param from_addr_str: 发件人友好名称
    :param password: 发件人邮件密码
    :param to_address: 收件人地址，格式为字符串，以逗号隔开
    :param header: 主题内容
    :return: 无返回值
    """
    # TODO 获取发件人的邮箱、发件内容处理（无人员处理），生成正文内容body
    body = ''
    # TODO 邮件发送
    send(smtp_server, smtp_port, from_addr, from_addr_str, password, to_address, header, body,)
    pass


# 主程序
def main():
    #  初始化数据库连接
    stone = stoneobject()
    # 初始化数据库中的4个基本表
    remove(stone)
    # 数据提取
    create_table(stone)
    targettimestr = input('请输入定点时间，例如8:00')
    targettime = datetime.time(int(targettimestr.split(':')[0]), int(targettimestr.split(':')[1]))
    while True:
        if datetime.datetime.now().hour == targettime.hour:
            date = datetime.date.today() + datetime.timedelta(days=0)
            # TODO 判断时间，如果时间为月末最后一天，调用unloading 存到 月转存表
            unloading(stone, today=date, afterday=3, number=7, table=WeekMapping)
            email_send()
            # TODO 判断时间，如果时间为周五，调用unloading 存到 周转存表
            unloading(stone, today=date, afterday=3, number=30, table=MonthMapping)
            email_send()
            # TODO 数据清理函数
            time.sleep(timer(targettime))
        else:
            print(timer(targettime))
            time.sleep(timer(targettime))
        pass
    pass


if __name__ == '__main__':
    main()
