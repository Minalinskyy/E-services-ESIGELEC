'''
Created on 8 août 2017

@author: Wentong ZHANG
'''
from builtins import str
    
def connectEsv(writeFile = False, usr='', pwd=''):
    import sys
    import time
    from selenium import webdriver
    from selenium.webdriver import DesiredCapabilities
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException

    
    print('download abs starts')
    loginUrl = 'https://cas.esigelec.fr/cas/login?service=http%3A%2F%2Fe-services.esigelec.fr%2Fj_spring_cas_security_check'
    '''r = requests.get(loginUrl)'''
    
    
    # PhantomJS setting
    headers = { 'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        #'Accept-Encoding':'gzip,deflate','
        'Accept-Language':'fr-FR;q=0.8',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.2; WOW64; rv:47.0) Gecko/20100101 Firefox/47.0',
        'Connection': 'keep-alive'
    }
    
    desired_capabilities = DesiredCapabilities.PHANTOMJS.copy()
    
    for key, value in headers.items():
        desired_capabilities['phantomjs.page.customHeaders.{}'.format(key)] = value
    
    desired_capabilities['phantomjs.page.settings.userAgent'] = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.116 Safari/537.36'
    

    # start driver
    driver = webdriver.PhantomJS(desired_capabilities=desired_capabilities)
    # driver = webdriver.Chrome()
    driver.set_window_size(1920,1080) # default window size of phantomjs is too small that css may display diffrently from other browser
    
    # load url
    driver.get(loginUrl)
    assert "CAS" in driver.title
    
    # login
    username = driver.find_element_by_name('username')
    password = driver.find_element_by_name('password')
    
    if usr and pwd:
        username.send_keys(usr)
        password.send_keys(pwd)
    else:
        try:
            from PlanningGetter import param
            username.send_keys(param.username)
            password.send_keys(param.password)
        except:
            print('usr,pwd error')
            sys.exit("No username or password inputted")

    
    # load home page
    password.send_keys(Keys.RETURN)

    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.ID, "form:entree_6420842"))
        )
        driver.find_element_by_link_text('Scolarité').click()
        time.sleep(3)
        driver.find_element_by_link_text('Mes Absences').click()
        time.sleep(3)
        # driver.find_element_by_link_text('Mes absences en 2016-2017').click()
        driver.find_element_by_link_text('Mes absences en 2017-2018').click()
    except Exception as e:
        print('click abs page error')
        if isinstance(e, TimeoutException):
            sys.exit("E-services index 60s time out")
        else:
        	pass

    # Mes absences page wait
    try:
        WebDriverWait(driver, 60).until(
            EC.title_contains('Mes absences')
        )
    except Exception as e:
        driver.close()
        print('enter abs page error,driver closed')
        if isinstance(e, TimeoutException):
            sys.exit("Mes absences 60s time out")
        else:
        	pass
    
    assert "Mes absences en 2017-2018" in driver.title
    
    # Abs detail wait
    try:
        WebDriverWait(driver, 40).until(
            EC.invisibility_of_element_located((By.ID, 'form:j_idt18'))
        )
    except Exception as e:
        driver.close()
        print('abs page load window error,driver closed')
        if isinstance(e, TimeoutException):
            sys.exit("Absences detail 40s time out")
        else:
        	pass

    try:
        nbrabs = int(driver.find_element_by_id('form:nbrAbs').text)
        nbrpg = (nbrabs-1)//20 + 1

        abs = []

        for i in range(0, nbrpg):
            try:
                time.sleep(1)
                if(i > 0 and i < nbrpg): # exclude the first page, we need to click '>' to go to next page
                    driver.find_element_by_xpath("//a[@class='ui-paginator-next ui-state-default ui-corner-all']").click()
                time.sleep(1)
                # wait the load window disappears
                try:
                    WebDriverWait(driver, 40).until(
                        EC.invisibility_of_element_located((By.ID, 'form:j_idt18'))
                    )
                except Exception as e:
                    print('Next abs page load window error page '+str(i+1)+', driver closed')
                    driver.close()
                    if isinstance(e, TimeoutException):
                        sys.exit("Abs detail 40s time out in week " + str(i + 1))
                    else:
                    	pass

                # only get the web code which includes 2 numbers for the first time
                if (i == 0):
                    absDetail = WebDriverWait(driver, 40).until(
                        EC.visibility_of_element_located((By.ID, 'form:recapAbsence'))
                    )
                    abs.append(absDetail.get_attribute('innerHTML'))
                # wait the main table of abs loaded and save web code
                absDetail = WebDriverWait(driver, 40).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, 'ui-datatable-tablewrapper'))
                )
                abs.append(absDetail.get_attribute('innerHTML'))
            except Exception as e:
                driver.close()
                print('get abs info error, driver closed')
                if isinstance(e, TimeoutException):
                    sys.exit("AbsDetail 40s time out")
                else:
                	pass

    except:
        abs = []
        abs.append('zero abs at this moment')

    driver.close()
    
    print('download abs finishes')
    # write to file or return list
    if writeFile:
        import io
        with io.open('abs.txt','w',encoding='utf8') as f:
        	for item in abs:
        		f.write("%s[[[Magic]]]\n" % item)
    else:
        return abs

 
if __name__ == '__main__':
    connectEsv(writeFile = True,usr='w.zhang.12',pwd='ZWTzwt940331')
     
