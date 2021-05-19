#!/usr/bin/python
# FINAL SCRIPT updated as of 11th October 2019
# Task - Function for Automagica ARRPA
# Scriptors - RPA Team

# Declare Python libraries needed for this script
import sys
import pandas as pd
import datetime
from automagica import *
from connector.dbconfig import read_db_config
from connector.connector import PostgresConnector
from automagic.automagica_common import wait_for_overlag
#from utils.audit_trail import audit_log
#from utils.logging import logging
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

import os
#from pathlib import Path
#current_path = os.path.dirname(os.path.abspath(__file__))
#dti_path = str(Path(current_path).parent)

#Set the location for download
def set_profile(source):
  try:
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference( "browser.download.manager.showWhenStarting", False )
    profile.set_preference("browser.download.dir", source + "new\\")

    return profile
  except Exception as error:
    print(error)

def download_Jasper(browser, day, month, year):
  print('Downloading Jasper file..')
  TixCreatedDateFrom = '%s-01-01' % year
  TixCreatedDateTo = '%s-%s-%s' %(year,month,day)
  try:
    
    browser.find_element_by_xpath('/html/body/div[10]/div[2]/div[2]/div/div/div[2]/ul/div[1]/label/input').send_keys(TixCreatedDateFrom)
    browser.find_element_by_xpath('/html/body/div[10]/div[2]/div[2]/div/div/div[2]/ul/div[2]/label/input').send_keys(TixCreatedDateTo)
    Wait(2)
    browser.find_element_by_xpath('/html/body/div[10]/div[2]/div[3]/button[2]').click()
    Wait(20)
    browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div/div[1]/div[4]/ul/li[1]/ul/li[2]/button/span/span[2]').click()
    Wait(2)
    browser.find_element_by_xpath('/html/body/div[6]/div[1]/div/ul/li[3]/p').click()
    Wait(10)
    
  except Exception as error:
    print('An error occured during download')
    print(error)
    #browser.close()

# Function to open a web browser
def initJasper(source, username=None, password=None):
  #audit_log('Logging in Jasper')
  print('Logging in Jasper')

  try:
    profile = set_profile(source)
    # Open using FireFox    
    browser = webdriver.Firefox(executable_path = dti_path+r'/selenium_interface/geckodriver.exe', firefox_profile=profile)
    # Check if both username and password are not null
    if username != None and password != None:
      user=username
      pwd=password

    else:
      config = read_db_config(dti_path+r'\config.ini', 'jasper')
      user = config['user']
      pwd = config['password']

    #browser = webdriver.Firefox(executable_path = r'\\Dtisvr2\rpa_cba\RPA.DTI\selenium_interface\geckodriver.exe')
    browser.get('https://jasper.aa-international.com/jasperserver/login.html')
    browser.find_element_by_xpath('/html/body/div[4]/div/div/div[1]/form/div/div/div[2]/div[2]/fieldset[1]/input[1]').send_keys(user)
    browser.find_element_by_xpath('/html/body/div[4]/div/div/div[1]/form/div/div/div[2]/div[2]/fieldset[1]/input[2]').send_keys(pwd)
    Wait(2)
    browser.find_element_by_xpath('/html/body/div[4]/div/div/div[1]/form/div/div/div[3]/div/button').click()
    Wait(2)
    browser.get('https://jasper.aa-international.com/jasperserver/flow.html?_flowId=viewReportFlow&_flowId=viewReportFlow&ParentFolderUri=%2FReport%2FMAX_MY&reportUnit=%2FReport%2FMAX_MY%2FMY_Adhoc_Accounting_Report_v3_Excel&standAlone=true')
    Wait(15)
    return browser

  except Exception as error:
    print('An error occured during logging in')
    print(error)
    #browser.close()

#source = "C:\\Users\\SHAHRUL.RAIMIE\\Documents\\Finance\\New\\"
#now = datetime.datetime.now()
#year=now.year
#month=now.month
#day=now.day
#browser = initJasper(source)
#download_Jasper(browser, day, month, year)
#browser.close()
