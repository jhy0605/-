import smtplib
import pymysql
import datetime
import time
import openpyxl
from smtplib import SMTP_SSL
from email.mime.text import MIMEText  # 发送字符串的邮件
from email.mime.multipart import MIMEMultipart
from email.header import Header  # 邮件标题
from email.mime.application import MIMEApplication  # 用于添加附件
import os


class yinc_call_list:
    def __init__(self):
        self.pwd_phat = ''
        self.host = '10.10.100.81'
        self.port = 3306
        self.user = 'root'
        self.passwd = 'jinwan_88888'
        self.name = 'line_phone_record'
        self.datetime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.mysqldb = pymysql.connect(host=self.host, port=self.port, user=self.user, password=self.passwd,
                                       database=self.name)  # 打开数据库链接
        self.curs = self.mysqldb.cursor()
        self.file_path = os.path.join(self.pwd_phat, r'银川白名单通话记录{}.xlsx'.format(
            datetime.datetime.now().strftime("%Y%m%d%H%M%S")))

    # 读取号码清单txt文件
    def phone_list(self):
        phone_list_all = []
        p = open(os.path.join(self.pwd_phat, r'号码清单.txt'), 'r').readlines()
        for phone in p:
            phone = phone.replace('\n', '')
            phone_list_all.append(phone)
        return tuple(phone_list_all)

    # 导出通话记录存储为xlsx文件
    def call_list(self):
        # 获取当天的数据
        phone_list = yinc_call_list.phone_list(self)

        sql_x = "SELECT callType as '通话类型', callerNo as '主叫', calledNo as '被叫', startTime as '开始时间'," \
                " endTime as '结束时间' FROM line_record_download_down_call_list " \
                "WHERE DATE(startTime) = CURDATE() and displayNumber in {};".format(phone_list)

        # print(sql_x)
        print('正在导出......')
        self.curs.execute(sql_x)  # 执行sql语句
        fetchall = self.curs.fetchall()
        excel_taber = openpyxl.Workbook()  # 创建excel对象
        sheet1 = excel_taber.active
        sheet1.append(['通话类型', '主叫', '被叫', '开始时间', '结束时间'])
        for call in fetchall:
            # print(call)
            sheet1.append(list(call))
        excel_taber.save(self.file_path)
        print('导出完成')
        time.sleep(3)

    # 发送邮件
    def smtp_email(self):
        # 定义发送的服务器，账号密码，收件人
        host_server = 'smtp.tz.mail.wo.cn'  # smtp服务器
        sender = 'jianghongyu@jwcredit.cn'  # 发件人邮箱
        pwd = 'Ysyyrps.77'  # 授权码

        receiver = ['hongjianwu@jwcredit.cn']  # 收件人邮箱
        cc = ['jianghongyu@jwcredit.cn']  # 抄送邮箱

        # 构建邮件内容
        msg = MIMEMultipart()
        msg['Subject'] = Header('联通白名单卡通话记录', 'utf-8')  # 邮件标题
        msg['From'] = sender  # 发件人
        msg['To'] = ';'.join(receiver)  # 收件人
        msg['Cc'] = ';'.join(cc)  # 抄送

        # 定义邮件正文内容
        mail_content_txt = """您好：<br>
                           &nbsp;&nbsp;今日通话记录统计已完成，附件请查收，谢谢！
                           """
        # 添加签名
        sign = """<br/><br/><br/><hr width="20%" align="left"/>
                        <div style="font-family:楷体;font-size:10.5" >
                        <p>广东金湾信息科技有限公司</p>
                        <p>蒋宏宇&emsp;系统工程师</p>
                        <p>电话：18152745300</p>
                        <p>邮箱：jianghongyu@jwcredit.cn</p>
                        <p>地址：510660 广州市天河区珠吉路8号713</p>
                        </div>
                        <div style="font-family:Arial;font-size:11" >
                        <p>本邮件所包含的信息和文件为广东金湾信息科技有限公司（以下称“公司”）的财产，均属公司保密信息，且包含旨在为邮件收件人所用的保密和专属信息。如果您不是收件人，任何披露、传播、复制或基于此邮件的任何未经授权的行为均被严格禁止。如您误收该邮件，请立即销毁并通知发件人。公司不因本邮件之误发而放弃其所享之任何权利。如本邮件的保密信息包含个人信息，请遵守中华人民共和国法律法规规定的各项个人信息保护义务。</p>
                        <p>The information contained in this e-mail is the property and confidential information of Guangdong Jinwan Co., Ltd. (hereinafter referred as Jinwan) and contains confidential and privileged proprietary information intended only for the personal and confidential use of the individual(s) whom it is addressed. If you are not the intended recipient of this email, any disclosure, dissemination, copying or unauthorized use of this message is strictly prohibited. In the case of an error, please immediately destroy this message and kindly notify the sender. Jinwan do not waive confidentiality or privilege by the mistransmission of this e-mail. If any personal information is contained in this e-mail, please comply with the obligations of personal information protection of the laws and regulations of the People's Republic of China.</p>
                        </div>        
                """
        mail_content_txt += sign

        # 发送html内容
        msg.attach(MIMEText(mail_content_txt, 'html', 'utf-8'))

        # 发送带附件的邮件
        if os.path.exists(self.file_path):
            attachment = MIMEApplication(open(self.file_path, 'rb').read())
            attachment.add_header('Content-Disposition', 'attachment', filename=os.path.split(self.file_path)[1])
            msg.attach(attachment)

        # 登录服务器
        try:
            smtp = SMTP_SSL(host_server)  # ssl登录
            smtp.login(sender, pwd)  # 登录
            smtp.sendmail(sender, receiver + cc, msg.as_string())  # 发送邮件
            smtp.quit()  # 退出
            print('邮件发送成功')
        except smtplib.SMTPException:
            print('无法发送邮件')

        if os.path.exists(self.file_path):
            os.remove(self.file_path)


if __name__ == '__main__':
    ycl = yinc_call_list()
    # ycl.pwd_phat = r'E:\银川定时发送白名单卡通话记录'
    ycl.pwd_phat = r'D:\BaiduSyncdisk\python\develop\联通云录音实时下载\发送通话记录'
    ycl.call_list()  # 导出通话记录存储为xlsx文件
    ycl.smtp_email()  # 发送邮件py
