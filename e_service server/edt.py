'''
Created on 8 août 2017

@author: Wentong ZHANG
'''
from E_services import cours_download, AnalyseCours
from datetime import datetime
import pytz
import smtplib

def sendmail(type,body):

  fromaddr = "jerry574638690@gmail.com"
  toaddr = "jerry574638690@gmail.com"
  msg = "\r\n".join([
    "From: jerry574638690@gmail.com",
    "To: jerry574638690@gmail.com",
    "Subject: ESIGELEC COURS CHANGE",
    "",
    body
    ])
  server = smtplib.SMTP('smtp.gmail.com:587')
  server.ehlo()
  server.starttls()
  server.login(fromaddr, "574638690")
  server.sendmail(fromaddr, toaddr, msg)
  server.quit()
  print('send email ok -- '+ type)

print("Actual time:")
loc = pytz.utc.localize(datetime.utcnow())
paristime = loc.astimezone(pytz.timezone('Europe/Paris'))
fmt = '%Y-%m-%d %H:%M:%S %Z%z'
print(paristime.strftime(fmt)+'\n\n')

body = ''

try:
    rescours = AnalyseCours.analyse(inners=cours_download.connectEsv(writeFile=False, usr='w.zhang.12', pwd='ZWTzwt940331', nbWeeks = 8))
except:
    rescours = 'bug'

if (rescours == 'changed'):
    body = 'You have new cours or cours changed, check it if you want\n'
elif(rescours == 'bug'):
    body = 'Get cours information bugged, check your code\n'

if body != '':
  sendmail('cours', body)

print('result cours:'+ rescours+'\n\n')

loc = pytz.utc.localize(datetime.utcnow())
paristime = loc.astimezone(pytz.timezone('Europe/Paris'))
print(paristime.strftime(fmt)+'\n\n')