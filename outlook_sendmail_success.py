import smtplib
from string import Template
from pathlib import Path
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase #用於承載附檔
from email import encoders #用於附檔編碼

# 開啟 CSV 檔案
all_content = []
with open('XXX.csv', newline='', encoding='utf-8') as csvfile:
  # 讀取 CSV 檔案內容
  rows = csv.reader(csvfile)
  # 以迴圈輸出每一列
  for row in rows:
    all_content.append(row)
all_content = all_content[1:]
for row in all_content:
    username = "XXX@oitc.com.tw"
    password = "XXX"
    mail_from = "XXX@oitc.com.tw"
    mail_to = row[1]
    mail_subject = '信件主題'
    mimemsg = MIMEMultipart()
    mimemsg['From']=mail_from
    mimemsg['To']=mail_to
    mimemsg['Subject']=mail_subject
    # 以下是可以加入多個附加檔案
    attachments = ["XXX.ics", "fileXXX.txt"]
    for file in attachments:
        with open(file, 'rb') as fp:
            add_file = MIMEBase('application', "octet-stream")
            add_file.set_payload(fp.read())
        encoders.encode_base64(add_file)
        add_file.add_header('Content-Disposition', 'attachment', filename=file)
        mimemsg.attach(add_file)
    template = Template(Path("remind_mail.html").read_text("utf-8"))
    body = template.substitute({"name": row[0], "subject": row[2], "time": row[3]})
    mimemsg.attach(MIMEText(body, 'html'))
    connection = smtplib.SMTP(host='smtp.office365.com', port=587)
    connection.starttls()
    connection.login(username,password)
    connection.send_message(mimemsg)
    connection.quit()
    print("已寄出")

