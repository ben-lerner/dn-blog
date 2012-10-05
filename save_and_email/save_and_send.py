#Save & email an Excel workbook
#Python 2.7 code, free to use
import smtplib
import os
import tempfile
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email import Encoders

#sends an email through a SMTP server [GMail in this example]
#using standard Python libraries
def sendThruGmail(to, subject, text, attachmentPath):
    #your username & password on your email server
    gmail_user = ""
    gmail_pwd = ""
    
    #main content of the email
    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = to
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    #attachment
    part = MIMEBase("application", "octet-stream")
    part.set_payload(open(attachmentPath, "rb").read())
    Encoders.encode_base64(part)
    (dirpath, filename) = os.path.split(attachmentPath)
    part.add_header("Content-Disposition", "attachment; filename=\"" + filename + "\" ")
    msg.attach(part)

    #GMail specific authentication
    mailServer = smtplib.SMTP("smtp.gmail.com", 587)
    mailServer.ehlo()
    mailServer.starttls()
    mailServer.ehlo()
    mailServer.login(gmail_user, gmail_pwd)

    #actual send
    mailServer.sendmail(gmail_user, to, msg.as_string())
    mailServer.close()

#save a copy to a temporary file (doesn't save/overwrite current the original file)
filename = tempfile.gettempdir() + "\\my_workbook.xlsx"
print "Saving a spare copy to " + filename
save_copy(filename)
subject = "Excel workbook"
body = "See the attachment. \n\n\nSent from my DataNitro"
sendThruGmail("to@email.com", subject, body, filename)
print "Done!"

