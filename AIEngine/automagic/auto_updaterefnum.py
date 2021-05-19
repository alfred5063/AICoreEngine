#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - MSIG

# Declare Python libraries needed for this script
import sys
import pandas as pd
import time
import os
import wx
from utils.notification import send
from automagica import *
from connector.dbconfig import read_db_config
from selenium import webdriver
from pathlib import Path
from utils.audit_trail import audit_log
from utils.logging import logging
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)

# Function to open a web browser
def initBrowser(main_df, msig_base):

  global browser

  try:
    #msig_base.username = ''
    #msig_base.password = ''
    # Check if both username and password are not null
    if msig_base.username != '' and msig_base.password != '':
      # Open using FireFox
      browser = webdriver.Firefox(executable_path = dti_path+r'/selenium_interface/geckodriver.exe')
      browser.maximize_window()
      #browser.get('https://marcmy.asia-assistance.com/marcmy/')
      browser.get('http://marcmyuat.aan.com:8080/marcmy/')
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').clear()
      Wait(3)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').send_keys(msig_base.username)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').clear()
      Wait(3)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').send_keys(msig_base.password)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/button').click()
      Wait(3)

    else:
      config = read_db_config(dti_path+r'\config.ini', 'marc-finance')
      dti_user = config['user']
      dti_pass = config['password']

      # Open using FireFox
      browser = webdriver.Firefox(executable_path = dti_path+r'/selenium_interface/geckodriver.exe')
      browser.maximize_window()
      #browser.get('https://marcmy.asia-assistance.com/marcmy/')
      browser.get('http://marcmyuat.aan.com:8080/marcmy/')
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').clear()
      Wait(3)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').send_keys(dti_user)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').clear()
      Wait(3)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').send_keys(dti_pass)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/button').click()
      Wait(3)

    checkresultdf = msig_dataRetrieve(main_df, msig_base)

    close_Marc(browser)
    return checkresultdf

  except Exception as error:
    close_Marc(browser)
    logging("MSIG - initBrowser", error, msig_base)
    print(error)

# Update Reference Number in MARC
def msig_dataRetrieve(main_df, msig_base):

  admNoList = main_df['Adm. No']
  clientRefNumList = main_df['Client Ref.Num']
  refnum_updated = list()
  refnum_notupdated = list()
  i = 0
  for i in range(len(admNoList)):
    try:
      msig_marc(int(admNoList[i]), str(clientRefNumList[i]), msig_base)
      time.sleep(5)
    except Exception:
      audit_log("MSIG - Admission No %s can't be processed. Please check manually" %int(admNoList[i]), "Completed...", msig_base)
      continue
  for i in range(len(admNoList)):
    try:
      refnum_checkdf = refnum_check(int(admNoList[i]), msig_base,refnum_updated,refnum_notupdated)
      time.sleep(5)
    except Exception as error:
      audit_log("MSIG - Reference number checking", error, msig_base)
      continue
  return refnum_checkdf
# Menuevering dataset based on Case ID and update the record
def msig_marc(admNoList, clientRefNumList, msig_base):

    # Direct to Inpatient Cases search page
    #browser.get('https://marcmy.asia-assistance.com/marcmy/pages/inpatient/cases/inpatient_cases_search.xhtml')
    browser.get('http://marcmyuat.aan.com:8080/marcmy/pages/inpatient/cases/inpatient_cases_search.xhtml')
    Wait(5)

    # Find client's information using Admission ID
    browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[1]/td[4]/input').send_keys(admNoList)
    Wait(5)
    browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[11]/td[2]/button[1]').click()
    Wait(5)

    # Click client's name (a link)
    browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[4]/a').click()
    Wait(5)

    # Click update button
    browser.find_element_by_xpath('/html/body/div[1]/div[3]/form[1]/div[1]/div[1]/div[1]/a[5]').click()
    Wait(5)

    # Clear the textarea and paste new reference number
    try:
      element_textare = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[1]/table/tbody/tr[1]/td[2]/input'
      textare = browser.find_element_by_xpath(element_textare)
      textare.clear()
      ActionChains(browser).move_to_element(textare).click(textare).send_keys(clientRefNumList).perform()
    except Exception:
      try:
        time.sleep(3)
        element_textare = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[1]/table/tbody/tr[1]/td[2]/input'
        textare = browser.find_element_by_xpath(element_textare)
        textare.click()
        textare.clear()
        textare.send_keys(clientRefNumList)
      except Exception as error:
        logging("MSIG - Automagica - Clear the textarea and paste new reference number", error, msig_base)
      pass
    Wait(5)


    # Save the record
    browser.find_element_by_xpath('/html/body/div[1]/div[3]/form[1]/div[1]/div[1]/div[1]/a[5]').click()
    Wait(5)

    # Unlock the Case ID
    try:
      ActionChains(browser).send_keys(Keys.RIGHT).perform()
      element_unlock = '/html/body/div[1]/div[3]/form[1]/div[1]/div[1]/div[2]/span[5]'
      my_unlock = browser.find_element_by_xpath(element_unlock)
      my_unlock.click()
      Wait(5)
      audit_log("MSIG - Case ID [ %s ] has been unlocked." % admNoList, "Completed...", msig_base)
    except Exception:
      try:
        ActionChains(browser).send_keys(Keys.RIGHT).perform()
        element_unlock = '/html/body/div[1]/div[3]/form[1]/div[1]/div[1]/div[2]/a[3]/img'
        my_unlock = browser.find_element_by_xpath(element_unlock)
        my_unlock.click()
        audit_log("MSIG - Case ID [ %s ] has been unlocked." % admNoList, "Completed...", msig_base)
      except Exception as error:
        logging("MSIG - Automagica - Unlock the Case ID properly", error, msig_base)
      pass
    Wait(5)

def refnum_check(admNoList, msig_base,refnum_updated,refnum_notupdated):
  #browser.get('https://marcmy.asia-assistance.com/marcmy/pages/inpatient/cases/inpatient_cases_search.xhtml')
  browser.get('http://marcmyuat.aan.com:8080/marcmy/pages/inpatient/cases/inpatient_cases_search.xhtml')
  Wait(5)
  # Find client's information using Admission ID
  browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[1]/td[4]/input').send_keys(admNoList)
  Wait(5)
  browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[11]/td[2]/button[1]').click()
  Wait(5)

  # Click client's name (a link)
  browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[4]/a').click()
  Wait(5)

  # Check id reference number is empty
  try:
    element_textare = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[1]/table/tbody/tr[1]/td[2]/input'
    textare = browser.find_element_by_xpath(element_textare)
    refnum_value = textare.get_attribute('value')
    if refnum_value != '':
      print("Case ID [ %s ] is updated." % admNoList)
      refnum_updated.append(admNoList)
      audit_log("MSIG - Case ID [ %s ] is updated." % admNoList, "Completed...", msig_base)
    else:
      print("Case ID [ %s ] is not updated." % admNoList)
      refnum_notupdated.append(admNoList)
      audit_log("MSIG - Case ID [ %s ] is not updated." % admNoList, "Completed...", msig_base)
  except Exception:
    #print("checking refnum error")
    logging("MSIG - Automagica - Reference number value checking", error, msig_base)
    pass
  Wait(5)
  refnum_updateddf = pd.DataFrame(refnum_updated, columns = ['Updated Case ID'])
  refnum_notupdateddf = pd.DataFrame(refnum_notupdated, columns = ['Not Updated Case ID'])
  refnum_checkdf = pd.concat([refnum_updateddf, refnum_notupdateddf], axis=1)
  return refnum_checkdf
  

# Logout from Marc and close browser
def close_Marc(browser):

  try:

    browser.close()

  except Exception as error:
    print(error)
