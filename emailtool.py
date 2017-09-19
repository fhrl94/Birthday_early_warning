import smtplib
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr


def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

def send(smtp_server, smtp_port, from_addr, from_addr_str, password, to_address, header, body,):
    """

    :param smtp_server: SMTP地址
    :param smtp_port: SMTP端口
    :param from_addr: 发件人邮箱
    :param from_addr_str: 发件人友好名称
    :param password: 发件人邮件密码
    :param to_address: 收件人地址，格式为字符串，以逗号隔开
    :param header: 主题内容
    :param body: 正文内容
    :return: 无返回值
    """
    # smtp_server =r'smtp.qq.com'
    # 企业邮箱
    # smtp_server = r'smtp.qiye.163.com'
    # smtp_server = r'smtp.163.com'

    # 正文
    msg = MIMEText(body, 'html', 'utf-8')
    # 主题，
    msg['Subject'] = Header(header, 'utf-8').encode()
    # 发件人别名
    msg['From'] = _format_addr('{name}<{addr}>'.format(name=from_addr_str, addr=from_addr))
    # 收件人别名
    msg['To'] = to_address

    # server = smtplib.SMTP(smtp_server, 25)
    # server.login(from_addr, password)
    server = smtplib.SMTP_SSL(smtp_server, smtp_port)
    # QQ SSL 端口 587
    # 网易 SSL端口 994、465
    # server.starttls()
    server.login(from_addr, password)
    # server.set_debuglevel(1)
    server.sendmail(from_addr, to_address.split(','), msg.as_string())
    server.quit()
    print('发送成功')
