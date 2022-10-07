import win32com.client as win32
from string import Template
from pathlib import Path


# 寄信測試
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'XXXu@oitc.com.tw'
mail.Subject = '主題'
mail.Body = 'hello'
template = Template(Path(r".\html_template\remind_mail.html").read_text())
body = template.substitute({"name": "JI","content1":"KO"})
mail.HTMLBody = body
# To attach a file to the email (optional):
attachment  = "Path to the attachment"
# mail.Attachments.Add(attachment)
mail.Send()

