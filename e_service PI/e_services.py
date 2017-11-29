'''
Created on 8 août 2017

@author: Wentong ZHANG
'''
from E_services import abs_download, AnalyseAbs, notes_download, AnalyseNotes, cours_download, AnalyseCours

body = ''
try:
    resabs = AnalyseAbs.analyse(inners=abs_download.connectEsv(writeFile=False, usr='w.zhang.12', pwd='???'))
except:
    resabs = 'bug'
if (resabs == 'changed'):
    body = 'You have new abs, check it if you want'

try:
    resnotes = AnalyseNotes.analyse(inners=notes_download.connectEsv(writeFile=False, usr='w.zhang.12', pwd='???'))
except:
    resnotes = 'bug'

if (resnotes == 'changed'):
    body = 'You have new notes, check it if you want'

try:
    rescours = AnalyseCours.analyse(
        inners=cours_download.connectEsv(writeFile=False, usr='w.zhang.12', pwd='???', nbWeeks=10))
except:
    pass

if body != '':
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    fromaddr = "jerry574638690@gmail.com"
    toaddr = "jerry574638690@gmail.com"
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = "CHANGEMENT OF E-SERVICES"

    msg.attach(MIMEText(body, 'plain'))

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, "password of your gmail")
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()