# -*- coding: UTF-8 -*-
import poplib
import base64
import email
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr


class get_command:

    def __init__(self):
        self.useraccount = 'lmm920316@163.com'
        self.password = 'LCBLWGXFHJNAZLTM'
        self.pop3_server = 'pop.163.com'
        self.msg = []

    def link_server(self):
        server = poplib.POP3(self.pop3_server)
        server.user(self.useraccount)
        server.pass_(self.password)
        # email_num, email_size = server.stat()# 获取邮箱状态。消息数量，消息总大小
        rsp, msg_list, rsp_siz = server.list()  # 请求消息列表 rsp:邮件数量、总大小，msg_list：消息编号和大小，rsp_size：消息大小
        rsp, msglines, msgsiz = server.retr(len(msg_list))  # 获得最新的一个邮件
        # server.dele(len(msg_list))# 删除这封邮件
        server.quit()  # 关闭与服务器的连接，释放资源
        msg_content = b'\r\n'.join(msglines).decode('utf-8')  # 编码一下，形成一个字符串
        msg = Parser().parsestr(text=msg_content)  # 组成一个EmailMassage类
        self.msg = msg

    def parser_subject(self, msg):  # 用来解析邮件主题
        subject = msg['Subject']
        value, charset = decode_header(subject)[0]
        if charset:
            value = value.decode(charset)
        print('邮件主题： {0}'.format(value))

    def parser_address(self, msg):  # 用来解析邮件来源
        hdr, addr = parseaddr(msg['From'])
        name, charset = decode_header(hdr)[0]  # name 发送人邮箱名称， addr 发送人邮箱地址
        if charset:
            name = name.decode(charset)
        print('发送人邮箱名称: {0}，发送人邮箱地址: {1}'.format(name, addr))

    def parser_content(self, msg):  # 解析邮件内容
        content = msg.get_payload()
        # 文本信息
        content_charset = content[0].get_content_charset()  # 获取编码格式
        text = content[0].as_string().split('base64')[-1]
        text_content = base64.b64decode(text).decode(content_charset)  # base64解码

        # 添加了HTML代码的信息
        content_charset = content[1].get_content_charset()
        text = content[1].as_string().split('base64')[-1]
        html_content = base64.b64decode(text).decode(content_charset)

        print('文本信息: {0}\n添加了HTML代码的信息: {1}'.format(text_content, html_content))

    def decode_str(self, s):  # 字符编码转换
        value, charset = decode_header(s)[0]
        if charset:
            value = value.decode(charset)
        return value

    def get_att(self, msg):
        attachment_files = []
        for part in msg.walk():
            file_name = part.get_filename()  # 获取附件名称类型
            contType = part.get_content_type()

            if file_name:
                h = email.header.Header(file_name)
                dh = email.header.decode_header(h)  # 对附件名称进行解码
                filename = dh[0][0]
                if dh[0][1]:
                    filename = self.decode_str(str(filename, dh[0][1]))  # 将附件名称可读化
                    print(filename)
                data = part.get_payload(decode=True)  # 下载附件
                att_file = open('D:\\数模作业\\' + filename, 'wb')  # 在指定目录下创建文件，注意二进制文件需要用wb模式打开
                attachment_files.append(filename)
                att_file.write(data)  # 保存附件
                att_file.close()
        return attachment_files


if __name__ == '__main__':
    my_email = get_command()
    my_email.link_server()
    my_email.parser_subject(my_email.msg)
