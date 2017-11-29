'''
Created on 8 août 2017

@author: Wentong ZHANG
'''
from E_services import abs_download, AnalyseAbs, notes_download, AnalyseNotes
from datetime import datetime
import pytz
import smtplib

def sendmail(type,body):

  fromaddr = "jerry574638690@gmail.com"
  toaddr = "jerry574638690@gmail.com"
  msg = "\r\n".join([
    "From: jerry574638690@gmail.com",
    "To: jerry574638690@gmail.com",
    "Subject: ESIGELEC NOTE/ABS CHANGE",
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
    resabs = AnalyseAbs.analyse(inners=abs_download.connectEsv(writeFile=False, usr='w.zhang.12', pwd='ZWTzwt940331'))
except:
    resabs = 'bug'


if (resabs == 'changed'):
    body = 'You have new abs, check it if you want\n'
elif(resabs == 'bug'):
    body = 'Get abs information bugged, check your code\n'

if body != '':
  sendmail('abs', body)

print('result abs:'+resabs+'\n\n')
body = ''

try:
    resnotes = AnalyseNotes.analyse(inners=notes_download.connectEsv(writeFile=False, usr='w.zhang.12', pwd='ZWTzwt940331'))
except:
    resnotes = 'bug'

if (resnotes == 'changed'):
    body = 'You have new notes, check it if you want\n'
elif(resnotes == 'bug'):
    body = 'Get notes information bugged, check your code\n'

if body != '':
  sendmail('notes', body)

print('result notes:'+resnotes+'\n\n')

loc = pytz.utc.localize(datetime.utcnow())
paristime = loc.astimezone(pytz.timezone('Europe/Paris'))
print(paristime.strftime(fmt)+'\n\n')