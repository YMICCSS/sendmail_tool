import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template
from pathlib import Path
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase #用於承載附檔
from email import encoders #用於附檔編碼

# 以下可測試信件發送
try:
    username = "XXXX.com.tw"
    password = "XXXX"
    mail_from = "XXXX.com.tw"
    mail_to = "XXXX.com.tw"
    mail_subject = 'XXXX 會前提醒'
    # 寄信物件初始化
    mimemsg = MIMEMultipart()
    mimemsg['From']=mail_from
    mimemsg['To']=mail_to
    mimemsg['Subject']=mail_subject
    # 加入模板變數
    template = Template(Path("remind_mail.html").read_text("utf-8"))
    body = template.substitute({"name": "1", "subject": "2", "time": "3"})
    mimemsg.attach(MIMEText(body, 'html'))
    # 伺服器連接
    connection = smtplib.SMTP(host='mail.centurydev.com.tw', port=25)
    connection.ehlo()
    connection.login(username,password)
    connection.send_message(mimemsg)
    connection.quit()
except Exception as e:
    print("python 發生錯誤 ",e)


