'''
Created on 8 août 2017

@author: Wentong ZHANG
'''
from builtins import str
    
def connectEsv(writeFile=False, usr='', pwd='', nbWeeks=8):
    import sys
    import time
    from selenium import webdriver
    from selenium.webdriver import DesiredCapabilities
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException
        
    print('download cours start')
    loginUrl = 'https://cas.esigelec.fr/cas/login?service=http%3A%2F%2Fe-services.esigelec.fr%2Fj_spring_cas_security_check'
    '''r = requests.get(loginUrl)'''
    
    # PhantomJS setting
    headers = { 'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        # 'Accept-Encoding':'gzip,deflate','
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
    #driver = webdriver.Firefox()
    driver.set_window_size(1920, 1080)  # default window size of phantomjs is too small that css may display diffrently from other browser
    
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
        print('usr,pwd error')
        sys.exit("No username or password inputted")

    
    # load home page
    password.send_keys(Keys.RETURN)
    try:
        assert "Page d'accueil" in driver.title
    except Exception as e:
        pass

    # home page wait, enter planning page
    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.ID, "form:entree_6420842"))
        )
        
        driver.find_element_by_link_text('Scolarité').click()
        time.sleep(3)
        driver.find_element_by_link_text('Mon planning').click()
    except Exception as e:
        print('click timetable page error')
        if isinstance(e, TimeoutException):
            sys.exit("E-services index 60s time out")
        else:
            pass
    
    # Mon Planning page wait
    try:
        WebDriverWait(driver, 60).until(
            EC.title_contains('Mon planning')
        )
    except Exception as e:
        driver.close()
        print('enter abs page error,driver closed')
        if isinstance(e, TimeoutException):
            sys.exit("Mon Planning 60s time out")
        else:
            pass
    
    assert "Mon planning" in driver.title
    
    # Planning detail wait
    try:
        WebDriverWait(driver, 40).until(
            EC.invisibility_of_element_located((By.XPATH, "//div[contains(@class, 'ui-dialog ui-widget ui-widget-content ui-corner-all ui-shadow ui-hidden-container ui-draggable ui-resizable')]"))
        )
    except Exception as e:
        driver.close()
        if isinstance(e, TimeoutException):
            sys.exit("Planning detail 40s time out")
        else:
            pass
    
    # skip void planning
    btn = driver.find_element_by_xpath("//button[@class='fc-next-button ui-button ui-state-default ui-corner-left ui-corner-right']")
    
    inners = []
    
    # get 8 weeks' lessons
    for i in range(0, nbWeeks):
        if i != 0:
            btn.click()
        # wait 'loading' window disappears
        try:
            WebDriverWait(driver, 40).until(
                EC.invisibility_of_element_located((By.XPATH, "//div[contains(@class, 'ui-dialog ui-widget ui-widget-content ui-corner-all ui-shadow ui-hidden-container ui-draggable ui-resizable')]"))
            )
        except Exception as e:
            driver.close()
            if isinstance(e, TimeoutException):
                sys.exit("Planning detail 40s time out in week " + str(i + 1))
            else:
                pass
                
        # get all lessons block
        blocks = driver.find_elements_by_xpath("//a[contains(@class, 'fc-time-grid-event fc-v-event fc-event fc-start fc-end')]")
    
        # click in all blocks and get lessons detail
        for block in blocks:
            block.click()
            
            try:
                time.sleep(1)
                modaleDetail = WebDriverWait(driver, 40).until(
                    EC.visibility_of_element_located((By.ID, 'form:modaleDetail'))
                )
                inners.append(modaleDetail.get_attribute('innerHTML'))
                driver.find_elements_by_xpath("//a[@class='ui-dialog-titlebar-icon ui-dialog-titlebar-close ui-corner-all']")[0].click()
            except Exception as e:
                driver.close()
                if isinstance(e, TimeoutException):
                    sys.exit("ModaleDetail 40s time out")
                else:
                    pass
    
    print('download cours finish')
    driver.close()
    
    
    # write to file or return list
    if writeFile:
        import io
        with io.open('cours.txt', 'w', encoding='utf8') as f:
            for item in inners:
                f.write("%s[[[Magic]]]\n" % item)
    else:
        return inners

    
if __name__ == '__main__':
    connectEsv(writeFile = True,usr='w.zhang.12',pwd='ZWTzwt940331',nbWeeks=3)    
    
