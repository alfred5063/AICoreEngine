#!/usr/bin/python
# FINAL SCRIPT updated as of 20th April 2020
# Workflow - MARC Report Download

# Declare Python libraries needed for this script
import time
import datetime
import os
import shutil
import selenium
from automagica import *
from selenium import webdriver
from pathlib import Path
from selenium.webdriver.common.action_chains import ActionChains
from directory.files_listing_operation import getListOfFiles

# Common path
#current_path = os.path.dirname(os.path.abspath(__file__))
#dti_path = str(Path(current_path).parent)

#config = read_db_config(dti_path+r'\config.ini', 'marc-finance')
dti_user = 'alfred.simbun'
dti_pass = 'M@rcpass4'

def previous_week_range(date):
  start_date = date + datetime.timedelta(-date.weekday(), weeks=-1)
  end_date = date + datetime.timedelta(-date.weekday() - 1)
  return start_date, end_date

def marc_autodownload(dti_user, dti_pass):
  global browser
  profile = webdriver.FirefoxProfile()
  profile.set_preference("browser.download.folderList", 2)
  profile.set_preference("browser.download.manager.showWhenStarting",False)
  profile.set_preference("browser.download.downloadDir", "\\dtisvr2\Finance_UAT\CML\MARC AP Report Weekly Download")
  profile.set_preference("browser.download.useDownloadDir", True)
  profile.set_preference("browser.download.forbid_open_with", False)
  profile.set_preference("browser.download.panel.shown", False)
  profile.set_preference("browser.helperApps.alwaysAsk.force", False)
  profile.set_preference('browser.helperApps.neverAsk.saveToDisk', ('application/vnd.ms-excel'))
  profile.update_preferences()

  browser = webdriver.Firefox(firefox_profile = profile, executable_path = r'C:/Users/ALFRED.SIMBUN/source/repos/rpa_engine/RPA.DTI/selenium_interface/geckodriver.exe')
  browser.maximize_window()
  #browser.get("about:config")
  browser.get('http://192.168.2.173:8080/marcmy/')
  browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').clear()
  browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').send_keys(dti_user)
  Wait(2)
  browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').clear()
  browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').send_keys(dti_pass)
  Wait(2)
  browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/button').click()

  # Direct to MY ACCPAC AP Report page
  browser.get('http://192.168.2.173:8080/marcmy/pages/report/my/my_inpt_import_purchase_report.xhtml')
  Wait(3)

  # Set the date ranges
  named_tuple = time.localtime()
  now_date = time.strftime("%d/%m/%Y", named_tuple)
  curr_startdate, curr_enddate = previous_week_range(datetime.date(named_tuple.tm_year, named_tuple.tm_mon, named_tuple.tm_mday))

  start_date = '/html/body/div[1]/div[3]/form/table/tbody/tr[1]/td[2]/span[1]/input'
  end_date = '/html/body/div[1]/div[3]/form/table/tbody/tr[1]/td[2]/span[2]/input'

  print("Set the start date as %s" % curr_startdate.strftime('%d/%m/%Y'))
  sDate = browser.find_element_by_xpath(start_date)
  ActionChains(browser).move_to_element(sDate).click(sDate).send_keys("%s" % curr_startdate.strftime('%d/%m/%Y')).perform()
  Wait(3)

  print("Set the end date as %s" % curr_enddate.strftime('%d/%m/%Y'))
  eDate = browser.find_element_by_xpath(end_date)
  ActionChains(browser).move_to_element(eDate).click(eDate).send_keys("%s" % curr_enddate.strftime('%d/%m/%Y')).perform()
  Wait(3)

  get_ready = '/html/body/div[1]/div[3]/form/h3'
  gReady = browser.find_element_by_xpath(get_ready)
  ActionChains(browser).move_to_element(gReady).click(gReady).perform()
  ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
  Wait(3)

  # Download the reports
  print("Download the AP report")
  download_button = '/html/body/div[1]/div[3]/form/table/tbody/tr[3]/td[2]/button[2]'
  dButton = browser.find_element_by_xpath(download_button)
  ActionChains(browser).move_to_element(dButton).click(dButton).perform()
  Wait(10)

  # Close MARC
  ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
  ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
  browser.find_element_by_xpath(get_ready).click()
  Wait(3)

  ActionChains(browser).send_keys(Keys.RIGHT).perform()
  close_element = '/html/body/div[1]/div[1]/div[2]/form/a[2]'
  closeButton = browser.find_element_by_xpath(close_element)
  ActionChains(browser).move_to_element(closeButton).click(closeButton).perform()

  browser.close()

  # Move and Rename the File
  source = r'C:\Users\ALFRED.SIMBUN\Downloads'
  destination = r'\\dtisvr2\Finance_UAT\CML\MARC AP Report Weekly Download'
  filelist = list(getListOfFiles(source))
  mylist = list(filter(lambda x: '.xls' in x, filelist))
  f = 0
  for f in range(len(mylist)):
    if 'accpac ap' in mylist[f].lower():
      new_filename = "\\MARC_AP_Report_" + str(curr_startdate) + "_to_" + str(curr_enddate) + ".xls"
      os.rename(mylist[f], source + new_filename)
      shutil.move(source + new_filename, destination)


marc_autodownload(dti_user, dti_pass)
