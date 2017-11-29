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

        
    print('download notes starts')
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
        driver.find_element_by_link_text('Mes notes').click()
    except Exception as e:
        print('click notes page error')
        if isinstance(e, TimeoutException):
            sys.exit("E-services index 60s time out")
        else:
            pass

    # Mes notes page wait
    try:
        WebDriverWait(driver, 60).until(
            EC.title_contains('Mes notes')
        )
    except Exception as e:
        driver.close()
        print('enter notes page error,driver closed')
        if isinstance(e, TimeoutException):
            sys.exit("Mes notes 60s time out")
        else:
            pass
    
    assert "Mes notes" in driver.title
    
    # Notes detail wait
    try:
        WebDriverWait(driver, 40).until(
            EC.invisibility_of_element_located((By.ID, 'form:j_idt18'))
        )
    except Exception as e:
        driver.close()
        print('notes page load window error,driver closed')
        if isinstance(e, TimeoutException):
            sys.exit("Notes detail 40s time out")
        else:
            pass

    notes = []
    click = True

    # start to get codes for notes of all pages
    while click:
        try:
            time.sleep(2)
            # wait notes page loaded and save web code
            notesDetail = WebDriverWait(driver, 40).until(
                EC.visibility_of_element_located((By.CLASS_NAME, 'ui-datatable-tablewrapper'))
            )
            notes.append(notesDetail.get_attribute('innerHTML'))
            # try to find this next page button, if found, that means there is still next page
            try:
                nextpage = driver.find_element_by_xpath("//a[@class='ui-paginator-next ui-state-default ui-corner-all']")
            except Exception as ee:
                click = False
            time.sleep(1)
            if(click == True):
                nextpage.click()
                try:
                    WebDriverWait(driver, 40).until(
                        EC.invisibility_of_element_located((By.ID, 'form:j_idt18'))
                    )
                except Exception as eee:
                    print('Next notes page load window error page '+str(i+1)+', driver closed')
                    driver.close()
                    if isinstance(e, TimeoutException):
                        sys.exit("Notes detail 40s time out in week " + str(i + 1))
                    else:
                        pass
        except Exception as e:
            print('get notes page error, driver closed')
            driver.close()
            click = False
            if isinstance(e, TimeoutException):
                sys.exit("NotesDetail 40s time out")
            else:
                pass

    driver.close()
    
    print('download notes finishes')
    # write to file or return list
    if writeFile:
        import io
        with io.open('notes.txt','w',encoding='utf8') as f:
        	for item in notes:
        		f.write("%s[[[Magic]]]\n" % item)
    else:
        return notes

 
if __name__ == '__main__':
    connectEsv(writeFile = True,usr='w.zhang.12',pwd='ZWTzwt940331')
     