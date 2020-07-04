# import os
# import win32com.client as win32
# import datetime
# from testFile import getPathInfo,readConfig
#
#
# read_conf = readConfig.ReadConfig()
# subject = read_conf.get_email('subject')#从配置文件中读取，邮件主题
# app = str(read_conf.get_email('app'))#从配置文件中读取，邮件类型
# addressee = read_conf.get_email('addressee')#从配置文件中读取，邮件收件人
# cc = read_conf.get_email('cc')#从配置文件中读取，邮件抄送人
# mail_path = os.path.join(getpathInfo.get_Path(), 'result', 'report.html')#获取测试报告路径
#
# class send_email():
#     def outlook(self):
#         olook = win32.Dispatch("%s.Application" % app)
#         mail = olook.CreateItem(win32.constants.olMailItem)
#         mail.To = addressee # 收件人
#         mail.CC = cc # 抄送
#         mail.Subject = str(datetime.datetime.now())[0:19]+'%s' %subject#邮件主题
#         mail.Attachments.Add(mail_path, 1, 1, "myFile")
#         content = """
#                     执行测试中……
#                     测试已完成！！
#                     生成报告中……
#                     报告已生成……
#                     报告已邮件发送！！
#                     """
#         mail.Body = content
#         mail.Send()
#
#
# if __name__ == '__main__':# 运营此文件来验证写的send_email是否正确
#     print(subject)
#     send_email().outlook()
#     print("send email ok!!!!!!!!!!")


#两种方式，第一种是用的win32com,因为系统等各方面原因，反馈win32问题较多，建议改成下面的smtplib方式
import os
import smtplib
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class SendEmail(object):
    def __init__(self, username, passwd, recv, title, content,
                 file=None, ssl=False,
                 email_host='smtp.qq.com', port=25, ssl_port=465):
        self.username = username  # 用户名
        self.passwd = passwd  # 密码
        self.recv = recv  # 收件人，多个要传list ['a@qq.com','b@qq.com]
        self.title = title  # 邮件标题
        self.content = content  # 邮件正文
        self.file = file  # 附件路径，如果不在当前目录下，要写绝对路径
        self.email_host = email_host  # smtp服务器地址
        self.port = port  # 普通端口
        self.ssl = ssl  # 是否安全链接
        self.ssl_port = ssl_port  # 安全链接端口

    def send_email(self):
        msg = MIMEMultipart()
        # 发送内容的对象
        if self.file:  # 处理附件的
            file_name = os.path.split(self.file)[-1]  # 只取文件名，不取路径
            try:
                f = open(self.file, 'rb').read()
            except Exception as e:
                raise Exception('附件打不开！！！！')
            else:
                att = MIMEText(f, "base64", "utf-8")
                att["Content-Type"] = 'application/octet-stream'
                # base64.b64encode(file_name.encode()).decode()
                new_file_name = '=?utf-8?b?' + base64.b64encode(file_name.encode()).decode() + '?='
                # 这里是处理文件名为中文名的，必须这么写
                att["Content-Disposition"] = 'attachment; filename="%s"' % (new_file_name)
                msg.attach(att)
        msg.attach(MIMEText(self.content))  # 邮件正文的内容
        msg['Subject'] = self.title  # 邮件主题
        msg['From'] = self.username  # 发送者账号
        msg['To'] = ','.join(self.recv)  # 接收者账号列表
        if self.ssl:
            self.smtp = smtplib.SMTP_SSL(self.email_host, port=self.ssl_port)
        else:
            self.smtp = smtplib.SMTP(self.email_host, port=self.port)
        # 发送邮件服务器的对象
        self.smtp.login(self.username, self.passwd)
        try:
            self.smtp.sendmail(self.username, self.recv, msg.as_string())
            pass
        except Exception as e:
            print('出错了。。', e)
        else:
            print('发送成功！')
        self.smtp.quit()


if __name__ == '__main__':
    m = SendEmail(
        username='605686114@qq.com',
        passwd='tknthdkskckubefa',
        recv=['605686114@qq.com'],
        title='66',
        content='666',
        file=r'D:\testtest.txt',
        ssl=True,
    )
    m.send_email()
# import smtplib
# from email.mime.text import MIMEText
# from email.header import Header
#
# # 第三方 SMTP 服务
# mail_host = "smtp.qq.com"  # 设置服务器
# mail_user = "605686114@qq.com"  # 用户名
# mail_pass = "tknthdkskckubefa"  # 口令
#
# sender = '605686114@qq.com'
# receivers = ['605686114@qq.com']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱
#
# message = MIMEText('Python 邮件发送测试...', 'plain', 'utf-8')
# message['From'] = Header("菜鸟教程", 'utf-8')
# message['To'] = Header("测试", 'utf-8')
#
# subject = 'Python SMTP 邮件测试'
# message['Subject'] = Header(subject, 'utf-8')
#
# try:
#     smtpObj = smtplib.SMTP()
#     smtpObj.connect(mail_host, 25)  # 25 为 SMTP 端口号
#     smtpObj.login(mail_user, mail_pass)
#     smtpObj.sendmail(sender, receivers, message.as_string())
#     print( "邮件发送成功")
#
# except smtplib.SMTPException as e:
#     print( "Error: 无法发送邮件%s" %e)
