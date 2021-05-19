#!/usr/bin/python
# FINAL SCRIPT updated as of 9th Dec 2020
# Workflow - STP-DM

# Declare Python libraries needed for this script
import sys
import pandas as pd
import time
from datetime import datetime, timedelta
from automagica import *
from utils.Session import session
from connector.dbconfig import read_db_config
from utils.audit_trail import audit_log
from utils.logging import logging
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from automagic.automagica_sql_setting import *
from utils.application import *
from selenium.common.exceptions import TimeoutException
import os
from pathlib import Path


current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)

#base = stpdm_base
# Function to open a web browser
def initBrowser(source,base,username=None, password=None):
  #app = get_application_url('lonpac_login')
  profile = set_profile(source)
  try:
    if profile != None:
      browser = webdriver.Firefox(firefox_profile = profile, executable_path = dti_path + r'/selenium_interface/geckodriver.exe')
    else:
      browser = webdriver.Firefox(executable_path = dti_path + r'/selenium_interface/geckodriver.exe')
    if username != None and password != None:
      user = username
      pwd = password
    else:
      config = read_db_config(dti_path + r'\config.ini', 'LONPAC')
      user = config['user']
      pwd = config['password']
    browser.maximize_window()
    try:
      print("open browser")
      browser.get('https://www.elonpac.com/eCorporateTP/system/application/loginframe.jsp?pg=1')
      frame = browser.find_element_by_xpath('/html/frameset/frame[2]')
      browser.switch_to_frame(frame)
      browser.find_element_by_xpath('/html/body/form/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[2]/td/table/tbody/tr[1]/td[4]/input').clear()
      browser.find_element_by_xpath('/html/body/form/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[2]/td/table/tbody/tr[1]/td[4]/input').send_keys(user)
      time.sleep(1)
      browser.find_element_by_xpath('/html/body/form/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td[4]/input').clear()
      browser.find_element_by_xpath('/html/body/form/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td[4]/input').send_keys(pwd)
      time.sleep(1)
      browser.find_element_by_xpath('/html/body/form/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/input[2]').click()
      print(browser)
      return browser
    except Exception as error:
      print(error)
      logging("STP-DM - Username and Password LONPAC web portal.", error, base)
    return browser
  except Exception as error:
    return error

#Set the location for download
def set_profile(source):
  try:
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference( "browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.helperApps.neverAsk.openFile",
                           "text/plain,text/x-csv,text/csv,application/vnd.ms-excel,application/csv,application/x-csv,text/csv,text/comma-separated-values,text/x-comma-separated-values,text/tab-separated-values,application/pdf, application/zip")
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk",
                          "text/plain,text/x-csv,text/csv,application/vnd.ms-excel,application/csv,application/x-csv,text/csv,text/comma-separated-values,text/x-comma-separated-values,text/tab-separated-values,application/pdf, application/zip")
    profile.set_preference("browser.download.dir", source)
    return profile
  except Exception as error:
    return error

def date_extraction(start_date, end_date=None):

  try:
    start_date = datetime.strptime(start_date, '%d/%m/%Y')
    if end_date !=None:
      end_date = datetime.strptime(end_date, '%d/%m/%Y')
      dayfrom = start_date.strftime('%d')
      monthfrom = start_date.strftime('%m')
      yearfrom = start_date.strftime('%Y')
      dayto = end_date.strftime('%d')
      monthto = end_date.strftime('%m')
      yearto = end_date.strftime('%Y')
    elif end_date ==None:
      end_date = start_date + timedelta(days=6)
      dayfrom = start_date.strftime('%d')
      monthfrom = start_date.strftime('%m')
      yearfrom = start_date.strftime('%Y')
      dayto = end_date.strftime('%d')
      monthto = end_date.strftime('%m')
      yearto = end_date.strftime('%Y')
    return dayfrom, monthfrom, yearfrom, dayto, monthto, yearto
  except Exception as error:
    return error


def escape_alert(browser):
  try:
    WebDriverWait(browser, 3).until(EC.alert_is_present(),
                                   'Timed out waiting for PA creation ' +
                                   'confirmation popup to appear.')

    alert = browser.switch_to.alert
    alert.accept()
    return 'alert'
  except TimeoutException as error:
    return error




def download_lonpac(browser, base, datefrom, dateend=None):
  print('Downloading lonpac file..')
  now = datetime.now()
  if dateend == None:
    dayfrom, monthfrom, yearfrom, dayto, monthto, yearto = date_extraction(datefrom)
  elif dateend !=None :
    dayfrom, monthfrom, yearfrom, dayto, monthto, yearto = date_extraction(datefrom, dateend)
  try:
    browser.switch_to.default_content()
    app = get_application_url('lonpac_download')
    browser.get(app.application.appUrl)
    frame = browser.find_element_by_xpath('/html/frameset/frameset/frame')
    browser.switch_to_frame(frame)
    browser.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[6]/td[2]/a').click()
    browser.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[7]/td[2]/table/tbody/tr[2]/td[2]/a').click()
    browser.switch_to.default_content()
    frame2 = browser.find_element_by_xpath('/html/frameset/frameset/frameset/frame[1]')
    browser.switch_to_frame(frame2)
    time.sleep(1)
    selectdayfrom = Select(browser.find_element_by_xpath('/html/body/div/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select[1]'))
    selectdayfrom.select_by_value(dayfrom)
    alert1 = escape_alert(browser)
    if alert1 != '':
      selectmonthfrom = Select(browser.find_element_by_xpath('/html/body/div/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select[2]'))
      selectmonthfrom.select_by_value(monthfrom)
      alert2 = escape_alert(browser)

    if alert2 != '':
      selectyearfrom = Select(browser.find_element_by_xpath('/html/body/div/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select[3]'))
      selectyearfrom.select_by_value(yearfrom)
      alert3 = escape_alert(browser)

    if alert3 != '':
      selectdayto = Select(browser.find_element_by_xpath('/html/body/div/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select[4]'))
      selectdayto.select_by_value(dayto)
      alert4 = escape_alert(browser)

    if alert4 != '':
      selectmonthto = Select(browser.find_element_by_xpath('/html/body/div/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select[5]'))
      selectmonthto.select_by_value(monthto)
      alert5 = escape_alert(browser)


    if alert5 != '':
      selectyearto = Select(browser.find_element_by_xpath('/html/body/div/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select[6]'))
      selectyearto.select_by_value(yearto)
      alert6 = escape_alert(browser)


    if alert6 != '':
      browser.find_element_by_xpath('/html/body/div/form/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[5]/td/input').click()
      
      Flag = True
      print('Download lonpac completed...')
      return Flag
    
  except Exception as error:
    Flag = False
    logging("STP-DM - error while downloading.", error, base)
    print(error)
    return Flag


