import platform
import smtplib
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr

import os


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
    # 正文
    msg = MIMEText(body, 'html', 'utf-8')
    # 主题，
    msg['Subject'] = Header(header, 'utf-8').encode()
    # 发件人别名
    msg['From'] = _format_addr('{name}<{addr}>'.format(name=from_addr_str, addr=from_addr))
    # 收件人别名
    msg['To'] = to_address
    server = smtplib.SMTP_SSL(smtp_server, smtp_port)
    server.login(from_addr, password)
    server.sendmail(from_addr, to_address.split(','), msg.as_string())
    server.quit()
    print('发送成功')

def sendMultimedia(smtp_server, smtp_port, from_addr, from_addr_str, password, to_address, header, body, file):
    """

    :param smtp_server: SMTP地址
    :param smtp_port: SMTP端口
    :param from_addr: 发件人邮箱
    :param from_addr_str: 发件人友好名称
    :param password: 发件人邮件密码
    :param to_address: 收件人地址，格式为字符串，以逗号隔开
    :param header: 主题内容
    :param body: 正文内容
    :param file: 附件路径名称
    :return: 无返回值
    """
    # 正文
    msg = MIMEMultipart()
    # 主题，
    msg['Subject'] = Header(header, 'utf-8').encode()
    # 发件人别名
    msg['From'] = _format_addr('{name}<{addr}>'.format(name=from_addr_str, addr=from_addr))
    # 收件人别名
    msg['To'] = to_address
    msg.attach(MIMEText(body, 'html', 'utf-8'))
    with open(file, 'rb') as f:
        # 设置附件的MIME和文件名，这里是png类型:
        mime = MIMEBase("application", "zip")
        # 加上必要的头信息:
        #  widnows平台使用gbk ,无法全平台通用
        if platform.system() == 'Windows':
            mime.add_header('Content-Disposition', 'attachment', filename=('gbk', '', os.path.basename(file)))
        else:
            mime.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', os.path.basename(file)))
        # mime.add_header('Content-ID', '<0>')
        # mime.add_header('X-Attachment-Id', '0')
        # 把附件的内容读进来:
        mime.set_payload(f.read())
        # 用Base64编码:
        encoders.encode_base64(mime)
        # 添加到MIMEMultipart:
        msg.attach(mime)
    server = smtplib.SMTP_SSL(smtp_server, smtp_port)
    server.login(from_addr, password)
    server.sendmail(from_addr, to_address.split(','), msg.as_string())
    server.quit()
    print('发送成功')
