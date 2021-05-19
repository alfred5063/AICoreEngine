#!/usr/bin/python
# FINAL SCRIPT updated as of 29th January 2020
# Workflow - Finance SOA
# Version 1

# Declare Python libraries needed for this script
import sys
import pandas as pd
import time
from automagica import *
from utils.Session import session
from connector.dbconfig import read_db_config
from utils.audit_trail import audit_log
from utils.logging import logging
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from automagic.automagica_sql_setting import *
from utils.application import *
import os
from pathlib import Path

current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)

def setFirefoxPreference(download_folder):
  """
    browser.download.folderList tells it not to use default Downloads directory
    browser.download.manager.showWhenStarting turns of showing download progress
    browser.download.dir sets the directory for downloads
    browser.helperApps.neverAsk.saveToDisk tells Firefox to automatically download the files of the selected mime-types
    """
  profile = webdriver.FirefoxProfile()
  profile.set_preference('browser.download.folderList', 2) # custom location
  profile.set_preference('browser.download.manager.showWhenStarting', False)
  profile.set_preference("browser.download.useDownloadDir", True)
  profile.set_preference('browser.download.dir', download_folder)
  profile.set_preference("browser.download.manager.alertOnEXEOpen", False)
  profile.set_preference("browser.download.manager.closeWhenDone", False)
  profile.set_preference("browser.download.manager.focusWhenStarting", False)
  profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/octet-stream;charset=UTF-8;application/vnd.ms-excel')#vnd.ms-excel
  profile.set_preference("print.always_print_silent", True)
  profile.set_preference("browser.download.forbid_open_with", True)
  profile.set_preference("print.show_print_progress", False)
  return 

# Function to open a web browser
def initBrowser(base, workflow, profile=None, username=None, password=None):
  app = get_application_url('marc_login')
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
      if workflow == 'SOA':
        config = read_db_config(dti_path+r'\config.ini', 'marc-finance')
        user = config['user']
        pwd = config['password']
      else:
        config = read_db_config(dti_path+r'\config.ini', 'marc2')
        user = config['user']
        pwd = config['password']

    #browser = webdriver.Firefox(executable_path = r'\\Dtisvr2\rpa_cba\RPA.DTI\selenium_interface\geckodriver.exe')
    browser.get(app.application.appUrl)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'username')).send_keys(user)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'password')).send_keys(pwd)
    Wait(5)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'login_button')).click()
    return browser

  except Exception as error:
    print(error)

def wait_to_load(browser,element):
  #wait for element to be show,suggested use in loading page for browser
  wait=WebDriverWait(browser,2)
  wait.until(EC.visibility_of_element_located((By.XPATH, element)))

def loading(browser):
  WebDriverWait(browser, 2).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[6]/div/div')))
  # then wait for the element to disappear
  WebDriverWait(browser, 30).until(EC.invisibility_of_element_located((By.XPATH, '/html/body/div[1]/div[6]/div/div')))
  Wait(1)

def wait_for_overlag(browser):
  #wait for the webpage to load the overlay, overlay is the delay that make mouse cannot click
  wait_overlag=WebDriverWait(browser,20)
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[6]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[22]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[1]/div[6]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[16]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html/body/div[1]/div[6]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH, "/html/body/div[7]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH, "/html/body/div[6]/div[13]")))

def search_Marc_inpatient(browser, base, caseID):
  try:
    app = get_application_url('marc_search_inpatient')
    browser.get(app.application.appUrl)
    #insert case ID from list, search, and view client
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'search_id')).send_keys(str(caseID))
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'search_button')).click()
    element=get_xpath_by_key(app.application.appId, 'member_entry')
    wait_to_load(browser,element)
    browser.find_element_by_xpath(element).click()
    wait_for_overlag(browser)
    return app
    
  except Exception as error:
    print(caseID+' cant be found in Inpatient')

def search_Marc_outpatient(browser, base, caseID):
  try:
    app = get_application_url('marc_search_outpatient')
    browser.get(app.application.appUrl)
    #insert case ID from list, search, and view client
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'search_id')).send_keys(str(caseID))
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'search_button')).click()
    element = get_xpath_by_key(app.application.appId, 'member_entry')
    wait_to_load(browser,element)
    browser.find_element_by_xpath(element).click()
    wait_for_overlag(browser)
  except Exception as error:
    print(caseID+' cant be found in Outpatient')

def search_obnum_inpatient(browser, base, caseID):
  try:
    app = get_application_url('marc_search_inpatient')
    browser.get(app.application.appUrl)
    #insert case ID from list, search, and view client
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'search_obnum')).send_keys(str(caseID))
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'search_button')).click()
    loading(browser)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'member_entry')).click()
    return app
    
  except Exception as error:
    #print(caseID+' cant be found in Inpatient')
    return -1

def search_obnum_outpatient(browser, base, caseID):
  try:
    app = get_application_url('marc_search_outpatient')
    browser.get(app.application.appUrl)
    #insert case ID from list, search, and view client
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'search_obnum')).send_keys(str(caseID))
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'search_button')).click()
    loading(browser)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'member_entry')).click()
    return app

  except Exception as error:
    #print(caseID+' cant be found in Outpatient')
    return -1
