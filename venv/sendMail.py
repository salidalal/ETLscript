import win32com.client as win32
import datetime


mailList= ('sdalal@ecitele.com','Tamar.Adler@ecitele.com','Alexander.Linchevsky@ecitele.com','Yuval.Hen@ecitele.com' )
#mailList= ('sdalal@ecitele.com','sdalal@ecitele.com')
now = datetime.datetime.now().strftime("%d.%m - %b.xls")
filename = "C:/Users/sdalal/OneDrive - ECI Telecom LTD/PycharmProjects/untitled/venv/logs/" + now

attachment = filename


def sendMail():
    for mailadd in mailList:
        print("1")
        print(mailadd)

        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Subject = 'Etl daily report'+now
        mail.Body = 'daily etl report'
        # To attach a file to the email (optional):
        mail.Attachments.Add(attachment)
        mail.To = str(mailadd)
        mail.Send()







'''



import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import datetime
email_user = 'etlreportsender@gmail.com'
email_password = 'Dala2017'
email_send = 'sdalal@ecitele.com'
subject = 'Etl daily report'+now

msg = MIMEMultipart()
msg['From'] = email_user
msg['To'] = email_send
msg['Subject'] = subject

body = 'Hi there, this is the daily etl report!'
msg.attach(MIMEText(body,'plain'))





part = MIMEBase('application','octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition',"attachment; filename= "+filename)

msg.attach(part)
text = msg.as_string()
server = smtplib.SMTP('smtp.gmail.com',587)
server.starttls()
server.login(email_user,email_password)


server.sendmail(email_user,email_send,text)
server.quit()'''