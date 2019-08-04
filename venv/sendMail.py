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



