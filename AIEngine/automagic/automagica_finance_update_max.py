#!/usr/bin/python
# FINAL SCRIPT updated as of 24th Nov 2020
# Task - Function for Automagica ARRPA Sage to Max
# Scriptors - RPA Team

# Declare Python libraries needed for this script
import sys
import time
import pandas as pd
import datetime
from automagica import *
from connector.dbconfig import read_db_config
from automagic.automagica_common import wait_for_overlag
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
from utils.Session import session
from regex import sub, findall
from datetime import datetime
import numpy as np
import os
from pathlib import Path
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)

def wait_to_click(browser, element):
  wait = WebDriverWait(browser, 20)
  element = wait.until(EC.element_to_be_clickable((By.XPATH, element)))
  return element

#Function to open max
def login_max(base):
  audit_log('Logging in Max', 'Completed...', base)
  try:
    profile = webdriver.FirefoxProfile()
    browser = webdriver.Firefox(executable_path = dti_path+r'/selenium_interface/geckodriver.exe',firefox_profile = profile)
    browser.maximize_window()
    if str(base.username) != '' and str(base.password) != '':
      user = base.username
      pwd = base.password
    else:
      config = read_db_config(dti_path + r'\config.ini', 'max')
      user = config['user']
      pwd = config['password']

    app = get_application_url('max_login')
    browser.get(app.application.appUrl)
    #browser.get('https://sys.aa-international.com/max/')

    try:
      browser.find_element_by_xpath('/html/body/div[1]/div/form/table/tbody/tr[3]/td[2]/input').clear()
      browser.find_element_by_xpath('/html/body/div[1]/div/form/table/tbody/tr[3]/td[2]/input').send_keys(user)
      time.sleep(5)
      browser.find_element_by_xpath('/html/body/div[1]/div/form/table/tbody/tr[4]/td[2]/input').clear()
      browser.find_element_by_xpath('/html/body/div[1]/div/form/table/tbody/tr[4]/td[2]/input').send_keys(pwd)
    except Exception as error:
      logging("Finance sage to max - Automagica - click username and password fields.", error, base)

    time.sleep(5)

    try:
      browser.find_element_by_xpath('/html/body/div[1]/div/form/table/tbody/tr[5]/td[2]/button').click()
    except Exception as error:
      logging("Finance sage to max - Automagica - click login_button field.", error, base)

    time.sleep(5)

    audit_log('Completed update in Max', 'Completed...', base)
    return browser
  except Exception as error:
    print('An error occured during logging in: {0}'.format(error))
    logging('Login max error', error, base)
    return error

#browser = browser_max
#tixid = tixid1
#costid = costingid1
#dateinv = dateinv1
#numinv = numinv1
#codecurnrc = codecurnrc
#amtduehc = amtduehc
#base = max_base

def update_max(browser, tixid, costid, dateinv, numinv, codecurnrc, amtduehc, base):
  
  app = get_application_url('max_find_tixId')
  browser.get(app.application.appUrl)
  #browser.get('https://sys.aa-international.com/max/pages/tix/tix_search.xhtml')

  try:
    str_invdate = str(dateinv)
    temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})T([0-9]{2}):([0-9]{2}):([0-9]{2}).([0-9]{9})$", str_invdate)
    if temp != []:
      temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2})T([0-9]{2}):([0-9]{2}):([0-9]{2}).([0-9]{9})$", str_invdate)[0]
      dateinv = "%s/%s/%s" % (temp[2], temp[1], temp[0])
    else:
      str_invdate = str(dateinv)
      temp = findall("^([0-9]{4})-([0-9]{2})-([0-9]{2}) ([0-9]{2}):([0-9]{2}):([0-9]{2})$", str_invdate)[0]
      dateinv = "%s/%s/%s" % (temp[2], temp[1], temp[0])
  except Exception as error:
    logging('Format date error', error, base)

  try:
    element_inpat = '/html/body/div[1]/div[3]/form/div[2]/table/tbody/tr[1]/td[2]/input'
    myinput = browser.find_element_by_xpath(element_inpat)
    ActionChains(browser).move_to_element(myinput).click(myinput).send_keys("%s" % str(tixid)).perform()
  except:
    try:
      time.sleep(5)
      element_inpat = '/html/body/div[1]/div[3]/form/div[2]/table/tbody/tr[1]/td[2]/input'
      myinput = browser.find_element_by_xpath(element_inpat)
      myinput.send_keys("%s" % str(tixid))
    except Exception as error:
      logging("Finance sage to max - Automagica - find_tixId field.", error, base)

  time.sleep(5)
  try:
    element_inpat = '/html/body/div[1]/div[3]/form/div[2]/table/tbody/tr[9]/td[2]/button[1]'
    myinput = browser.find_element_by_xpath(element_inpat)
    ActionChains(browser).move_to_element(myinput).click(myinput).perform()
  except Exception as error:
    time.sleep(5)
    element_inpat = '/html/body/div[1]/div[3]/form/div[2]/table/tbody/tr[9]/td[2]/button[1]'
    myinput = browser.find_element_by_xpath(element_inpat)
    myinput.click()
    logging("Finance sage to max - Automagica - search_tixId field.", error, base)

  time.sleep(5)
  try:
    element_checkexist = '/html/body/div[1]/div[3]/form/div[2]/div[3]/div/div/div[3]/table/tbody/tr'
    myresult = browser.find_element_by_xpath(element_checkexist).text
  except:
    time.sleep(5)
    element_checkexist = '/html/body/div[1]/div[3]/form/div[2]/div[3]/div/div/div[3]/table/tbody/tr'
    myresult = browser.find_element_by_xpath(element_checkexist).text

  records_status = []
  if ('No records found.' in str(myresult)):
    status = "No cost id found. Tix id is {0}".format(str(tixid))
    records_status.append(status)
    return records_status, browser

  if str(myresult) != 'No records found.':

    # Check if the current Tix ID is locked. Unlocked if it is locked.
    time.sleep(5)
    try:
      element_inpat = '/html/body/div[1]/div[3]/form/div[2]/div[3]/div/div/div[3]/table/tbody/tr/td[29]'
      lock = browser.find_element_by_xpath(element_inpat).text
    except Exception as error:
      logging("Finance sage to max - Automagica - lock/unlock tixid.", error, base)

    error = ''
    try:
      element_inpat = '/html/body/div[1]/div[3]/form/div[2]/div[3]/div/div/div[3]/table/tbody/tr/td[1]/a'
      myinput = browser.find_element_by_xpath(element_inpat)
      ActionChains(browser).move_to_element(myinput).click(myinput).perform()
    except Exception as error:
      logging("Finance sage to max - Automagica - click_tixId field.", error, base)

    time.sleep(5)
    if str(lock) != '':
      try:
        element_inpat = '/html/body/div[1]/div[3]/form/div[2]/div[2]/a'
        myinput = browser.find_element_by_xpath(element_inpat)
        ActionChains(browser).move_to_element(myinput).click(myinput).perform()
      except:
        try:
          time.sleep(5)
          browser.find_element_by_xpath('//*[@id="j_idt260"]').click()
        except Exception as error:
          logging("Finance sage to max - Automagica - unlock the tixid.", error, base)

    time.sleep(6)
    try:
      element_inpat = '/html/body/div[1]/div[3]/form/div[1]/div[1]/a[4]/img'
      myinput = browser.find_element_by_xpath(element_inpat)
      ActionChains(browser).move_to_element(myinput).click(myinput).perform()
    except Exception as error:
      logging("Finance sage to max - Automagica - edit_costing field.", error, base)

    time.sleep(6)
    try:
      element_inpat = '/html/body/div[1]/div[3]/form/div[4]/h4[1]'
      myinput = browser.find_element_by_xpath(element_inpat)
      myinput.click()
      time.sleep(2)
      ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
      ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
      ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
      ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
      ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
      ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
      ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
    except Exception as error:
      logging("Finance sage to max - Automagica - cannot page down.", error, base)

    time.sleep(10)
    try:
      element_inpat = '/html/body/div[1]/div[3]/form/div[4]/div[9]/ul/li[6]'
      myinput = browser.find_element_by_xpath(element_inpat)
      ActionChains(browser).move_to_element(myinput).click(myinput).perform()
    except:
      try:
        time.sleep(10)
        browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[4]/div[9]/ul/li[6]').click()
      except Exception as error:
        logging("Finance sage to max - Automagica - Costing Tab", error, base)

    time.sleep(5)
    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()

    time.sleep(5)
    costidlist = []
    elems = []
    foundcostid = []
    found_elems = []
    page = []
    page_elements = []
    try:
      numpages = len(browser.find_element_by_xpath('//*[@id="tabs:subTixCostingTable_paginator_top"]/span').text)
      for p in range(numpages):
        p+=1
        if p <= numpages:
          page_elements = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[2]/div[1]/span/a[1]'.replace('a[1]', 'a['+str(p)+']')
          browser.find_element_by_xpath(page_elements).click()
          time.sleep(5)
          elems = browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[2]/div[3]/table/tbody').find_elements_by_tag_name('tr')
          for elem in elems:
            cost_id = elem.get_attribute('data-rk')
            if str(costid) == cost_id:
              foundcostid = costidlist
              found_elems = elems
              page = page_elements
            costidlist.append(cost_id)
    except Exception as error:
      logging("Finance sage to max - Automagica - check_costing_id field", error, base) 

    time.sleep(5)
    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()

    time.sleep(3)
    try:
      element_inpat = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[2]/div[5]/a[1]/span'
      myinput = browser.find_element_by_xpath(element_inpat)
      ActionChains(browser).move_to_element(myinput).click(myinput).perform()
    except Exception as error:
      logging("Finance sage to max - Automagica - Go to page 1", error, base)

    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
    ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()

    time.sleep(3)
    if str(costid) in foundcostid:
      for i in range(len(foundcostid)):
        if i == int(foundcostid.index(str(costid))):

          try:
            myinput = browser.find_element_by_xpath(page)
            ActionChains(browser).move_to_element(myinput).click(myinput).perform()
          except Exception as error:
            logging("Finance sage to max - Automagica - Go to page 1", error, base)

          ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
          ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
          ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()

          time.sleep(5)
          try:
            a = str(foundcostid.index(str(costid)))
            element_inpat = '//*[@id="tabs:subTixCostingTable:10:j_idt2132"]'.replace(':10:', ':'+a+':')
            #for uat
            #element_inpat = '//*[@id="tabs:subTixCostingTable:10:j_idt2100"]'.replace(':10:', ':'+a+':')
            #print(element_inpat)
            myinput = browser.find_element_by_xpath(element_inpat)
            time.sleep(2)
            ActionChains(browser).move_to_element(myinput).click(myinput).perform()
            time.sleep(5)
          except Exception as error:
            logging("Finance sage to max - Automagica - click edit_cost_id field.", error, base)

        else:
          #print(i)
          continue

      time.sleep(5)
      try:
        mylist = Select(browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[2]/table/tbody/tr/td[3]/table/tbody/tr[5]/td[2]/select'))
        time.sleep(2)
        mylist.select_by_visible_text(str(codecurnrc))
        time.sleep(5)
      except:
        try:
          mylist = Select(browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[2]/table/tbody/tr/td[3]/table/tbody/tr[5]/td[2]/select'))
          time.sleep(2)
          mylist.select_by_visible_text(str(codecurnrc))
          time.sleep(5)
        except Exception as error:
          logging("Finance sage to max - Automagica - 'Cost to Client Curr'", error, base)

      time.sleep(5)
      try:
        element_inpat = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[2]/table/tbody/tr/td[3]/table/tbody/tr[6]/td[2]/input'
        myinput = browser.find_element_by_xpath(element_inpat)
        myinput.clear()
        time.sleep(5)
        ActionChains(browser).move_to_element(myinput).click(myinput).send_keys("%s" % str(amtduehc)).perform()
      except Exception as error:
        logging("Finance sage to max - Automagica - 'Cost Amt to Client'", error, base)

      time.sleep(3)
      try:
        browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[2]/table/tbody/tr/td[4]/table/tbody/tr[3]/td[2]/input').clear()
      except Exception as error:
        logging("Finance sage to max - Automagica - clearing inv_num field.", error, base)

      time.sleep(3)
      try:
        element_inpat = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[2]/table/tbody/tr/td[4]/table/tbody/tr[3]/td[2]/input'
        myinput = browser.find_element_by_xpath(element_inpat)
        ActionChains(browser).move_to_element(myinput).click(myinput).send_keys("%s" % numinv).perform()
      except Exception as error:
        logging("Finance sage to max - Automagica - sending numinv inv_num field.", error, base)

      time.sleep(3)
      try:
        browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[2]/table/tbody/tr/td[4]/table/tbody/tr[2]/td[2]/span/input').clear()
      except Exception as error:
        logging("Finance sage to max - Automagica - clearing inv_date field.", error, base)

      time.sleep(3)
      try:
        element_inpat = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[2]/table/tbody/tr/td[4]/table/tbody/tr[2]/td[2]/span/input'
        myinput = browser.find_element_by_xpath(element_inpat)
        ActionChains(browser).move_to_element(myinput).click(myinput).send_keys("%s" % dateinv).perform()
        browser.find_element_by_xpath('/html[1]/body[1]/div[1]/div[3]/form[1]/div[4]/div[9]/div[1]/div[6]/div[1]/div[3]/div[2]/span[1]/div[2]/div[2]').click()
      except Exception as error:
        logging("Finance sage to max - Automagica - sending dateinv inv_date field.", error, base)

      time.sleep(3)
      #try:
      #  element_inpat = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[2]/table/tbody/tr/td[3]/table/tbody/tr[6]/td[2]/input'
      #  myinput = browser.find_element_by_xpath(element_inpat)
      #  myinput.click()
      #except Exception as error:
      #  logging("Finance sage to max - Automagica - stabilize the window 1.", error, base)

      time.sleep(3)
      try:
        element_checkexist = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[2]/table/tbody/tr/td[4]/table/tbody/tr[6]/td[2]/input'
        tick_isbilled = browser.find_element_by_xpath(element_checkexist).get_attribute("checked")
        if(tick_isbilled != "true"):
          try:
            element_inpat = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[2]/table/tbody/tr/td[4]/table/tbody/tr[6]/td[2]/input'
            myinput = browser.find_element_by_xpath(element_inpat)
            ActionChains(browser).move_to_element(myinput).click(myinput).perform()
          except Exception as error:
            logging("Finance sage to max - Automagica - click tick_isbilled_true field.", error, base)
        else:
          print('Bill already checked previously.')
      except Exception as error:
        logging("Finance sage to max - Automagica - click tick_isbilled_true field.", error, base)

      time.sleep(3)
      try:
        element_inpat = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/div[2]/div[1]'
        myinput = browser.find_element_by_xpath(element_inpat)
        ActionChains(browser).move_to_element(myinput).click(myinput).perform()
      except Exception as error:
        logging("Finance sage to max - Automagica - pull up popped-up window.", error, base)

      time.sleep(5)
      try:
        element_inpat = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/div[3]/div[2]/span/table[1]/tbody/tr/td/table/tbody/tr[2]/td[2]/button[1]'
        myinput = browser.find_element_by_xpath(element_inpat)
        ActionChains(browser).move_to_element(myinput).click(myinput).perform()
      except Exception as error:
        logging("Finance sage to max - Automagica - click button_update_costing field.", error, base)

      time.sleep(10)
      ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
      ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()

      time.sleep(10)
      try:
        element_inpat = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/span[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/textarea'
        myinput = browser.find_element_by_xpath(element_inpat)
        ActionChains(browser).move_to_element(myinput).click(myinput).perform()
      except:
        try:
          time.sleep(5)
          element_inpat = '/html/body/div[1]/div[3]/form/div[4]/div[9]/div/div[6]/div/span[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/textarea'
          myinput = browser.find_element_by_xpath(element_inpat)
          ActionChains(browser).move_to_element(myinput).click(myinput).perform()
        except Exception as error:
          logging("Finance sage to max - Automagica - click to stabilize the window 2.", error, base)

      time.sleep(10)
      status = "Cost id {0} updated.".format(costid)
      records_status.append(status)
    else:
      status = "No cost id {0} found.".format(costid)
      records_status.append(status)

    try:
      element_inpat = '/html/body/div[1]/div[4]'
      myinput = browser.find_element_by_xpath(element_inpat)
      myinput.click()
      time.sleep(5)
    except Exception as error:
      logging("Finance sage to max - Automagica - failed to Page Up.", error, base)

    ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
    ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
    ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
    ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
    ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
    ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
    ActionChains(browser).send_keys(Keys.PAGE_UP).perform()

    time.sleep(10)
    try:
      element_inpat = '/html/body/div[1]/div[3]/form/div[1]/div[1]/a[4]'
      myinput = browser.find_element_by_xpath(element_inpat)
      ActionChains(browser).move_to_element(myinput).click(myinput).perform()
    except:
      try:
        time.sleep(10)
        element_inpat = '/html/body/div[1]/div[3]/form/div[1]/div[1]/a[4]/img'
        browser.find_element_by_xpath(element_inpat).click()
        time.sleep(10)
        browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/a[4]/img').click()
        time.sleep(10)
        browser.find_element_by_xpath('/html[1]/body[1]/div[1]/div[3]/form[1]/div[4]').click()
      except Exception as error:
        logging("Finance sage to max - Automagica - click click_save field.", error, base)
    time.sleep(10)
    return records_status, browser

def close_max(browser):
  browser.close()

def get_xpath_from_list(xpathlist, key):
  values =  xpathlist[xpathlist["key"]==key].xpath
  value = ""
  for item in values:
    value = item
  return value

def wait_for_overlag(browser):
  wait_overlag=WebDriverWait(browser,20)
  wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2626")))
  wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt103")))
  wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2588")))
  wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2315")))
  wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2810")))
  wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2599")))
  wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2868")))
  wait_overlag.until(EC.invisibility_of_element_located((By.CLASS_NAME,'ui-widget-overlay ui-dialog-mask')))
  wait_overlag.until(EC.invisibility_of_element_located((By.ID,"/html[1]/body[1]/div[6]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.ID,"/html[1]/body[1]/div[22]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[1]/div[6]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[16]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[4]/div[1]/div[1]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[4]/div[1]/div[1]")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"//div[@class='ui-growl-item']")))
  wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"//div[@class='three-quarters']")))

