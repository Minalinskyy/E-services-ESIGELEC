# E-services ESIGELEC

This is a little program which we can use instead of the e-services of ESIGELEC.

Because the google server is used by too many people, there is a problem of pushing data into server randomly. So I gave the version with and without it to let everyone choose.

Follow this tutorial step by step and we can get a program which can do:

1, For the windows or UNIX(Mac/Linux) version, we will have GUI. We can choice to read the data from localhost, or can choice to login and download the latest data of your timetable, absent and notes as you want, and then save them into the localhost file. And when the program update the timetable, it will automatically push the latest timetable into a google calendar which you can set in form of event. So we can check it on smartphone as we want.

2, For the PI version, this is for the people who have a Raspberry PI. We can put this programme into Raspberry PI and use crontab command to let it run automatically and periodically. When the programme runs, it can update all the 3 tables of data(timetable, absent and notes), save into localhost file and push the latest timetable into google calendar.  And if there is any change for absent and notes, it will send an email to a setted email adresse to give you this information.


Here is all the 3rd party packages we need to install for this program:
(language environment: python 3.X)
Necessary for all 3 versions:
beautifulsoup4,
certifi,
chardet,
et-xmlfile,
google-api-python-client,
httplib2,
idna,
jdcal,
lxml,
oauth2client,
openpyxl,
pyasn1,
pyasn1-modules,
PyJWT,
PySocks,
python-dateutil,
pytz,
requests,
rsa,
selenium,
six,
uritemplate,
urllib3,
all packages can be installed by : pip install packagename (for ex: pip install selenium)

PS: for the server version, we also need the package: icalendar

PS: if you have installed python 2 and 3 on same computer, be careful of the probleme of pip install and pip3 install

After the installation, we can start to do setup of our program.

Attention, the next step only for the version with google calendar:

{{

At first, follow this tutorial https://developers.google.com/google-apps/calendar/quickstart/python. This is in order to get the .json file. But remember to change one line of the exemple code: line21 SCOPES = 'https://www.googleapis.com/auth/calendar.readonly' we need to change the url into https://www.googleapis.com/auth/calendar. This is to get an .json file which we can use to read and WRITE with the google calendar. VERY IMPORTANT!

PS: if you have done the tutorial with .readonly and see this instruction, the solution is go to the defaul path of your computer:(normally it's C://Users//username for windows and home/user/~ for UNIX) There should be a hidden folder called .credentials, enter this folder and delete the file in it. And then re-run the exemple code but without .readonly in URL.

Don't forget copy the .json file into our project folder. Put it into the folder E_services.

}}

And then, we need to install the web browser driver phantomJS. http://phantomjs.org/download.html Download the version which you need from this URL and put the .exe file into your environment path. (For windows is C:\Windows\System32)

Then we need to do different things by the versions.

For the first 3 versions with google calendar:

On 62,69,138 line of the file Analysecours.py, there is a field called CalendarID, change it to your google calendar ID which you want to use. You can find this ID on the page of setup of your google calendar.

PS!!!!!: the default value i put in code is 'primary' which means the default calendar of your google account. If you use the default calendar for yourself, please remember to create a new one called like e-services, and at the end of the calendar setting page there is a calendar id like 215vrsdu75s3ogbsdh58q8uj8c@group.calendar.google.com , just replace the 'primary' by '215vrsdu75s3ogbsdh58q8uj8c@group.calendar.google.com'

Only for the Raspberry PI version:

1,On 35,36 line of e_services.py, change the email adresse into your adresse and the adresse you want to send to. On 46 line of this file, put your password of your gmail.

2,On 10,17,26 line of e_services.py, set the username and password of your account.

3, Don't forget the change I've written for all 3 versions.

OK, normally when you run the programme with command python e_services.py(or python3 e_services.py), everything should be okay.

Have a nice try.
