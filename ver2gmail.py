import smtplib
import pymysql as dbms

db=dbms.connect(host="localhost" ,user="root", password="08101997",database="email")
c=db.cursor()
c.execute('select Mail,Name from employee')
email=list(c.fetchall())

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from email.utils import COMMASPACE, formatdate

connection=smtplib.SMTP('smtp.gmail.com',587)
connection.ehlo()
connection.starttls()
print("started")
connection.login('durga.physio.97@gmail.com','08101997')

for i in email:
    mesg = "Dear "+i[1]+", \n\n     Choice greetings!"
    subject="Hello msg from Nagamani"
    body = "{}".format( mesg)
    Body_full="subject: {} \n\n {}".format(subject,mesg)
    msg = MIMEMultipart()
    msg['To'] = i[0]
    msg['Subject'] = "Hello msg from Nagamani"
    part2 = MIMEText(body)
    msg.attach(part2)

    file=i[1]+".docx"
    filename = file
    config = Path('D:'+filename)
    if config.is_file():
        print("attached")
        attachment = open("D:"+filename, "rb")
        if (file == filename):
            p = MIMEBase('application', 'octet-stream')
            p.set_payload((attachment).read())
            encoders.encode_base64(p)
            p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
            msg.attach(p)
            text = msg.as_string()
            connection.sendmail('durga.physio.97@gmail.com',i[0],text)
    else:
        connection.sendmail('durga.physio.97@gmail.com', i[0],Body_full)
        print("not attatched")
    print("success")
    print("--------------------------")
connection.quit()
