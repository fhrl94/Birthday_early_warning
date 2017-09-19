
import datetime
import time

from TimerTask import timer

# 数据初始化
# 1、从金蝶中获取职员表（职员信息（姓名、工号、出生日期）、职务、职位ID）
# 2、职位表（上下级关系）
from emailtool import send


def create_table():
    # TODO 将金蝶数据转存到sqlite 数据库中
    # TODO 转存（在职人员） 职员表 姓名、工号、出生日期、职位唯一ID（涉及到sqlite支不支持UID），需要处理 职务的获取（存在兼职的情况）
    # TODO 转存（存在的岗位） 职位表 当前职位ID，父级职位ID
    pass


# 数据转存
# 将指定日期期间的人员放入 转存表 的表头，通过迭代获取相应的上级
def unloading():
    # TODO 日期处理迭代（可用函数实现），将时间区域的人转存到 月转存表 或 周转存表
    # TODO 上级获取迭代（可用函数实现）
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
    body=''
    # TODO 邮件发送
    send(smtp_server, smtp_port, from_addr, from_addr_str, password, to_address, header, body,)
    pass


# 主程序
def main():
    # TODO 初始化数据库连接
    # 初始化数据库中的2个基本表
    create_table()
    Targettimestr = input('请输入定点时间，例如8:00')
    Targettime = datetime.time(int(Targettimestr.split(':')[0]), int(Targettimestr.split(':')[1]))
    while True:
        if datetime.datetime.now().hour == Targettime.hour:
            # TODO 判断时间，如果时间为月末最后一天，调用unloading 存到 月转存表
            unloading()
            email_send()
            # TODO 判断时间，如果时间为周五，调用unloading 存到 周转存表
            unloading()
            email_send()
            # TODO 数据清理函数
            time.sleep(timer(Targettime))
        else:
            print(timer(Targettime))
            time.sleep(timer(Targettime))
        pass
    pass


if __name__ == '__main__':
    main()
