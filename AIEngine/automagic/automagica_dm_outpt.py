# Declare Python libraries needed for this script
#!/usr/bin/python
# FINAL SCRIPT updated as of 9th April 2021

import sys
import pandas as pd
import time
from automagica import *
from connector.dbconfig import read_db_config
from utils.audit_trail import audit_log
from utils.logging import logging
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from automagic.automagica_sql_setting import *
from utils.application import *
from selenium.common import exceptions as se
import os
from pathlib import Path
import pyautogui
from automagic.marc import *

current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)


def initBrowser(url=None, profile=None, username=None, password=None):
  
  app = get_application_url('marc_login', url=url)
  
  print('Access application: {0}:'.format(app))
  try:
    if profile !=None:
      # Open using FireFox    
      browser = webdriver.Firefox(firefox_profile=profile, executable_path = dti_path+r'/selenium_interface/geckodriver.exe')
    else:
      browser = webdriver.Firefox(executable_path = dti_path+r'/selenium_interface/geckodriver.exe')
    #browser.manage().window().maximize()
    # Check if both username and password are not null
    if username != None and password != None:
      user=username
      pwd=password
    else:
      #config = read_db_config(dti_path+r'\config.ini', 'marc')
      config = read_db_config(dti_path+r'\config.ini', 'marc')
      user = config['user']
      pwd = config['password']
  
    browser.get(app.application.appUrl)
    
    #audit_log('Login marc','Trying to load the login page at {0}'.format(app.application.appUrl), base)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'username')).send_keys(user)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'password')).send_keys(pwd)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'login_button')).click()
    wait_to_load(browser, "/html[1]/body[1]/div[1]/div[3]/div[1]/div[1]/div[2]/form[1]/button[1]", 60)
    valid_login = False
    try:
      login_status = browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[1]").text()
      if login_status == "Login/Password incorrect.":
        #logging('Invalid login/password', 'login id: {0}'.format(user), base, status=mode)
        valid_login = False
    except Exception as error:
      if "Message: Unable to locate element: /html[1]/body[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[1]" in str(error): 
        print('login marc successful.')
        valid_login = True
        #audit_log('Import to MARC', 'Login to MARC successfully', base, status=mode)
      else:
        print('error: {0}'.format(error))

    return browser, valid_login
   
  except Exception as error:
    #logging('login issue due to invalid password', error, base)
    browser.close()




def upload_import_outpt(filepath, browser):
  try:
    print('import outpatient file into marc')
    browser.get('http://marcmyuat.aan.com:8080/marcmy/pages/dashboard/dashboard_inpatient.xhtml?faces-redirect=true')
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[1]/ul/li[5]/a').click()
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[1]/ul/li[5]/ul/li[2]/a').click()
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[1]/ul/li[5]/ul/li[2]/ul/li[2]/a').click()
    browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[1]/ul/li[5]/ul/li[2]/ul/li[2]/ul/li[2]/a').click()
    uploadElement = browser.find_element_by_xpath('//*[@id="j_idt49_input"]')
    uploadElement.send_keys(filepath)
    browser.find_element_by_xpath('/html/body/div[1]/div[3]/div[1]/form/div[2]/div[1]/button[1]').click()
    print('import outpatient file into marc completed...')
  except Exception as error:
    browser.close()
  return browser

#browser, valid_login = initBrowser(url=None, profile=None, username=None, password=None)
#filepath = r'\\dtisvr2\Finance_UAT\SOA\iyliaasyqin.abdmajid@asia-assistance.com\Result\Sunway SOA as at 30042020_01042021184803.xlsx'
#upload_import_outpt(filepath, browser)
