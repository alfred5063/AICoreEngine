#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - REQ/DCA

# Declare Python libraries needed for this script
import time
import datetime
import json
import os
import selenium
import pyautogui
import keyboard
from utils.logging import logging
from connector.dbconfig import *
from connector.connector import MSSqlConnector
from automagica import *
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from pathlib import Path
from utils.audit_trail import audit_log
from selenium.webdriver.common.action_chains import ActionChains
from pynput.keyboard import Key, Controller

# Common path
current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)

# Browser Management
def login_Marc(reqdca_base):
  try:

    global browser

    printingpath = r'\\Dtisvr2\cba_uat\Print\DCAREQ\New'
    profile = webdriver.FirefoxProfile()
    profile.set_preference('browser.download.folderList', 2) # custom location
    profile.set_preference('browser.download.manager.showWhenStarting', False)
    profile.set_preference("browser.download.useDownloadDir", True)
    profile.set_preference('browser.download.dir', printingpath)
    profile.set_preference("browser.download.manager.alertOnEXEOpen", False)
    profile.set_preference("browser.download.manager.closeWhenDone", False)
    profile.set_preference("browser.download.manager.focusWhenStarting", False)
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/octet-stream;charset=UTF-8;application/vnd.ms-excel')#vnd.ms-excel
    profile.set_preference("print.always_print_silent", True)
    profile.set_preference("browser.download.forbid_open_with", True)
    profile.set_preference("print.show_print_progress", False)

    # Check if both username and password are not null
    if reqdca_base.password != '':
      browser = webdriver.Firefox(firefox_profile = profile, executable_path = dti_path+r'/selenium_interface/geckodriver.exe')
      browser.maximize_window()
      #browser.get('https://marcmy.asia-assistance.com/marcmy/')
      browser.get('http://marcmyuat.aan.com:8080/marcmy/')
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').clear()
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').send_keys(reqdca_base.username)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').clear()
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').send_keys(reqdca_base.password)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/button').click()
      audit_log("REQ/DCA - Automagica - Accessing MARC using [ %s ]." % reqdca_base.username, "Completed...", reqdca_base)
    else:
      config = read_db_config(dti_path+r'\config.ini', 'marc-finance')
      dti_user = config['user']
      dti_pass = config['password']

      browser = webdriver.Firefox(firefox_profile = profile, executable_path = dti_path+r'/selenium_interface/geckodriver.exe')
      browser.maximize_window()
      #browser.get('https://marcmy.asia-assistance.com/marcmy/')
      browser.get('http://marcmyuat.aan.com:8080/marcmy/')
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').clear()
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').send_keys(dti_user)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').clear()
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').send_keys(dti_pass)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/button').click()
      audit_log("REQ/DCA - Automagica - Accessing MARC using [ %s ]." % dti_user, "Completed...", reqdca_base)
    return browser
  except Exception as error:
    audit_log("REQ/DCA - Web browser failed. MARC database cannot be accessed.", "Completed...", reqdca_base)
    logging("REQDCA - Automagica - Failed access to web browser.", error, reqdca_base)

def navigate_page(browser, *args):

    #navigate to any page, refer to URL for parameter
    #example : http://192.168.2.173:8080/marcmy/pages/dashboard/dashboard_inpatient.xhtml
    #navigate_page(browser, "dashboard", "dashboard_inpatient")

    #page = 'https://marcmy.asia-assistance.com/marcmy/pages'
    page = 'http://marcmyuat.aan.com:8080/marcmy/pages'
    for arg in args:
        page = page + '/' + arg
    page = page + '.xhtml'
    return browser.get(page)

def search_Marc_inpatient(browser, caseid, reqdca_base):
  try:
    element_inpat = '/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[1]/td[4]'
    myinput = browser.find_element_by_xpath(element_inpat)
    ActionChains(browser).move_to_element(myinput).click(myinput).send_keys("%s" % caseid).perform()

    element_but = '/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[11]/td[2]/button[1]'
    mybut = browser.find_element_by_xpath(element_but)
    ActionChains(browser).move_to_element(mybut).click(mybut).perform()
    Wait(3)

    element_checklocked = '/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[30]'
    mylocked = browser.find_element_by_xpath(element_checklocked).text
    if mylocked == "":
      element_link = '/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[4]/a'
      mylink = browser.find_element_by_xpath(element_link)
      ActionChains(browser).move_to_element(mylink).click(mylink).perform()
      Wait(5)
      audit_log("REQ/DCA - Automagica - Searched [ %s ] in MARC." % caseid, "Completed...", reqdca_base)
      flag = '1'
    else:
      audit_log("REQ/DCA - Automagica - Case ID [ %s ] is currently LOCKED in MARC." % caseid, "Completed...", reqdca_base)
      audit_log("REQ/DCA - Automagica - Skipping Case ID [ %s ]." % caseid, "Completed...", reqdca_base)
      flag = '0'
      pass
    return flag
  except Exception as error:
    logging("REQDCA - Automagica - Failed searched for Inpatient in MARC.", error, reqdca_base)

def search_Marc_outpatient(browser, subcaseid, reqdca_base):
  try:
    element_inpat = '/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[1]/td[4]/input'
    myinput = browser.find_element_by_xpath(element_inpat)
    ActionChains(browser).move_to_element(myinput).click(myinput).send_keys("%s" % subcaseid).perform()

    element_but = '/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[7]/td[2]/button[1]'
    mybut = browser.find_element_by_xpath(element_but)
    ActionChains(browser).move_to_element(mybut).click(mybut).perform()
    Wait(3)

    element_checklocked = '/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[16]'
    mylocked = browser.find_element_by_xpath(element_checklocked).text
    if mylocked == "":
      element_link = '/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[1]/a'
      mylink = browser.find_element_by_xpath(element_link)
      ActionChains(browser).move_to_element(mylink).click(mylink).perform()
      Wait(5)
      audit_log("REQ/DCA - Automagica - Searched [ %s ] in MARC." % subcaseid, "Completed...", reqdca_base)
      flag = '1'
    else:
      audit_log("REQ/DCA - Automagica - Case ID [ %s ] is currently LOCKED in MARC." % subcaseid, "Completed...", reqdca_base)
      audit_log("REQ/DCA - Automagica - Skipping Case ID [ %s ]." % subcaseid, "Completed...", reqdca_base)
      flag = '0'
      pass
    return flag
  except Exception as error:
    logging("REQDCA - Automagica - Failed searched for Outpatient in MARC.", error, reqdca_base)

def close_Marc(browser):
  browser.close()

def wait_to_load(browser, element):
  wait = WebDriverWait(browser, 10)
  element = wait.until(EC.visibility_of_element_located((By.XPATH, element)))
  return element

def wait_to_click(browser, element):
  wait = WebDriverWait(browser, 20)
  element = wait.until(EC.element_to_be_clickable((By.XPATH, element)))
  return element

# Create documents for inpatient and outpatient
def create_doc_inpatient(browser, attention, amount, remarks, HNS, doc_template, insurer, caseid, address, cln, current_date, current_time, reqdca_base):

  # Edit document of client
  try:
    time.sleep(10)
    element_edit = '/html/body/div[1]/div[3]/form[1]/div[1]/div[1]/div[1]/a[5]/img'
    myedit = wait_to_click(browser, element_edit)
    myedit.click()
  except Exception:
    try:
      element_edit = '/html/body/div[1]/div[3]/form[1]/div[1]/div[1]/div[1]/a[5]'
      myedit = wait_to_click(browser, element_edit)
      ActionChains(browser).move_to_element(myedit).click(myedit).perform()
    except Exception as error:
      logging("REQDCA - Automagica - Inpatient - Edit document of client", error, reqdca_base)
    pass

  time.sleep(30)

  try:
    element_doc = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/ul/li[6]'
    mydoc = wait_to_click(browser, element_doc)
    mydoc.click()
    audit_log("REQ/DCA - Automagica - Inpatient - Edit document of client.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_doc = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/ul/li[6]'
      mydoc = wait_to_click(browser, element_doc)
      ActionChains(browser).move_to_element(mydoc).click(mydoc).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Edit document of client.", "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Inpatient - Edit document of client", error, reqdca_base)
    pass

  time.sleep(5)

  # Select option in dropdown menu
  try:
    Select(browser.find_element_by_xpath('/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/span/table/tbody/tr[2]/td[2]/select')).select_by_visible_text(doc_template)
    Wait(2)
    Select(browser.find_element_by_xpath('/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/span/table/tbody/tr[3]/td[4]/select')).select_by_visible_text(insurer)
    Wait(2)
    browser.find_element_by_xpath('/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/span/table/tbody/tr[3]/td[2]/input').send_keys(attention)
    audit_log("REQ/DCA - Automagica - Inpatient - Select option in dropdown menu.", "Completed...", reqdca_base)
  except Exception as error:
    logging("REQDCA - Automagica - Inpatient - Select option in dropdown menu", error, reqdca_base)
    pass

  time.sleep(5)

  # Update LG amount
  try:
    element_lg = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/span/table/tbody/tr[2]/td[4]/input'
    mylgamt = wait_to_click(browser, element_lg)
    mylgamt.clear()
    ActionChains(browser).move_to_element(mylgamt).click(mylgamt).send_keys("%s" % amount).perform()
    audit_log("REQ/DCA - Automagica - Inpatient - Update LG amount.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_lg = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/span/table/tbody/tr[2]/td[4]/input'
      mylgamt = wait_to_click(browser, element_lg)
      mylgamt.clear()
      ActionChains(browser).move_to_element(mylgamt).click(mylgamt).send_keys("%s" % amount).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Update LG amount.", "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Inpatient - Update LG amount", error, reqdca_base)
    pass

  time.sleep(5)

  # Select "Send as"
  try:
    element_save = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/span/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input'
    mysave = wait_to_click(browser, element_save)
    mysave.click()
    audit_log("REQ/DCA - Automagica - Inpatient - Select Send As.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_save = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/span/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input'
      mysave = wait_to_click(browser, element_save)
      ActionChains(browser).move_to_element(mysave).click(mysave).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Select Send As.", "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Inpatient - Select Send As", error, reqdca_base)
    pass

  time.sleep(5)
  ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
  time.sleep(5)

  # Create Document
  try:
    element_createdoc = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/span/table/tbody/tr[10]/td/button[1]'
    mycreatedoc = wait_to_click(browser, element_createdoc)
    mycreatedoc.click()
    audit_log("REQ/DCA - Automagica - Inpatient - Create Document.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_createdoc = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/span/table/tbody/tr[10]/td/button[1]'
      mycreatedoc = wait_to_click(browser, element_createdoc)
      ActionChains(browser).move_to_element(mycreatedoc).click(mycreatedoc).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Create Document.", "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Inpatient - Create Document", error, reqdca_base)
    pass

  time.sleep(60)

  # Close the calender view
  try:
    element_closecal = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/div[2]/div[2]/div/div/table/tbody/tr[1]/td/input'
    closecal = wait_to_click(browser, element_closecal)
    closecal.click()
  except Exception:
    try:
      time.sleep(3)
      element_closecal = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/div[2]/div[2]/div/div/table/tbody/tr[1]/td/input'
      closecal = wait_to_click(browser, element_closecal)
      closecal.click()
    except Exception as error:
      logging("REQDCA - Automagica - Inpatient - Close the calender view", error, reqdca_base)
    pass

  time.sleep(5)

  # Set the time first
  try:
    named_tuple = time.localtime()
    current_time = time.strftime("%H%M", named_tuple)
    element_time = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/div[2]/div[2]/div/div/table/tbody/tr[1]/td/input'
    mycurtime = wait_to_click(browser, element_time)
    mycurtime.click()
    mycurtime.clear()
    mycurtime.send_keys(current_time)
    audit_log("REQ/DCA - Automagica - Inpatient - Set time.", "Completed...", reqdca_base)
  except Exception:
    try:
      time.sleep(3)
      mycurtime = wait_to_click(browser, element_time)
      mycurtime.click()
      mycurtime.clear()
      mycurtime.send_keys(current_time)
      audit_log("REQ/DCA - Automagica - Inpatient - Set time.", "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Inpatient - Set time", error, reqdca_base)
    pass

  time.sleep(5)

  # Set date
  try:
    named_tuple = time.localtime()
    current_date1 = time.strftime("%d/%m/%Y", named_tuple)
    element_date = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/div[2]/div[2]/div/div/table/tbody/tr[1]/td/span/input'
    mycurdate = wait_to_click(browser, element_date)
    mycurdate.clear()
    ActionChains(browser).move_to_element(mycurdate).click(mycurdate).send_keys(current_date1).perform()
    audit_log("REQ/DCA - Automagica - Inpatient - Set date.", "Completed...", reqdca_base)
  except Exception:
    try:
      time.sleep(3)
      mycurdate = wait_to_click(browser, element_date)
      mycurdate.clear()
      ActionChains(browser).move_to_element(mycurdate).click(mycurdate).send_keys(current_date1).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Set date.", "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Inpatient - Set date", error, reqdca_base)
    pass

  time.sleep(5)

  # Close the calender view
  try:
    element_closecal = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/div[2]/div[2]/div/div/table/tbody/tr[1]/td/input'
    closecal = wait_to_click(browser, element_closecal)
    closecal.click()
  except Exception:
    try:
      closecal = wait_to_click(browser, element_closecal)
      closecal.click()
    except Exception as error:
      logging("REQDCA - Automagica - Inpatient - calender view", error, reqdca_base)
    pass

  time.sleep(5)

  # Set the browser to popup window
  try:
    iframes = browser.find_elements_by_xpath("//iframe")
    browser.switch_to.frame(iframes[0])
    audit_log("REQ/DCA - Automagica - Inpatient - Set the browser to popup window.", "Completed...", reqdca_base)
  except Exception:
    try:
      iframes = browser.find_elements_by_xpath("//iframe")
      browser.switch_to.frame(iframes[0])
      audit_log("REQ/DCA - Automagica - Inpatient - Set the browser to popup window.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Set the browser to popup window")
      logging("REQDCA - Automagica - Inpatient - Set the browser to popup window", error, reqdca_base)
    pass

  time.sleep(5)

  # Edit 'bill to' field
  try:
    element_bill = '/html/body/div/h3[2]'
    mybill = wait_to_click(browser, element_bill)
    mybill.click()
    ActionChains(browser).move_to_element(mybill).click(mybill).send_keys("%s" % address).perform()
    audit_log("REQ/DCA - Automagica - Inpatient - Set the Bill To.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_bill = '/html/body/div/h3[2]'
      mybill = wait_to_click(browser, element_bill)
      mybill.click()
      ActionChains(browser).move_to_element(mybill).click(mybill).send_keys("%s" % address).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Set the Bill To.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Edit 'bill to' field")
      logging("REQDCA - Automagica - Inpatient - Set the Bill To", error, reqdca_base)
    pass

  time.sleep(5)

  # Inserting HNS running number
  try:
    element_hns = '/html/body/div/table[1]/tbody/tr[2]/td[3]'
    myhns = wait_to_click(browser, element_hns)
    ActionChains(browser).move_to_element(myhns).click(myhns).send_keys("%s" % str(HNS)).perform()
    audit_log("REQ/DCA - Automagica - Inpatient - Inserting HNS running number.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_hns = '/html/body/div/table[1]/tbody/tr[2]/td[3]'
      myhns = wait_to_click(browser, element_hns)
      ActionChains(browser).move_to_element(myhns).click(myhns).send_keys("%s" % str(HNS)).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Inserting HNS running number.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Inserting HNS running number")
      logging("REQDCA - Automagica - Inpatient - Inserting HNS running number", error, reqdca_base)
    pass

  time.sleep(5)

  # Inserting Client Listing Number
  try:
    element_cln = '/html/body/div/table[1]/tbody/tr[2]/td[4]'
    mycln = wait_to_click(browser, element_cln)
    ActionChains(browser).move_to_element(mycln).click(mycln).send_keys("%s" % str(cln)).perform()
    audit_log("REQ/DCA - Automagica - Inpatient - Inserting Client Listing Number.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_cln = '/html/body/div/table[1]/tbody/tr[2]/td[4]'
      mycln = wait_to_click(browser, element_cln)
      ActionChains(browser).move_to_element(mycln).click(mycln).send_keys("%s" % str(cln)).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Inserting Client Listing Number.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Inserting Client Listing number")
      logging("REQDCA - Automagica - Inpatient - Inserting Client Listing Number", error, reqdca_base)
    pass

  time.sleep(5)

  # Remove terms and conditions for Credit Note
  try:
    if int(float(amount)) < int(0):
      clear1 = '/html/body/div/p[1]'
      clear_line1 = browser.find_element_by_xpath(clear1)
      clear_line1.click()
      Wait(2)
      clear_line1.clear()
      Wait(2)
      clear2 = '/html/body/div/p[2]'
      clear_line2 = browser.find_element_by_xpath(clear2)
      ActionChains(browser).move_to_element(clear_line2).click(clear_line2).perform()
      Wait(2)
      clear_line2.clear()
      Wait(2)
      clear3 = '/html/body/div/p[1]'
      clear_line3 = browser.find_element_by_xpath(clear3)
      ActionChains(browser).move_to_element(clear_line3).click(clear_line3).send_keys(" ").perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Removed terms and conditions for Credit Note.", "Completed...", reqdca_base)
    else:
      audit_log("REQ/DCA - Automagica - Inpatient - Terms and conditions were not removed.", "Completed...", reqdca_base)
  except Exception:
    try:
      if int(float(amount)) < int(0):
        clear1 = '/html/body/div/p[1]'
        clear_line1 = browser.find_element_by_xpath(clear1)
        ActionChains(browser).move_to_element(clear_line1).click(clear_line1).perform()
        Wait(2)
        clear_line1.clear()
        Wait(2)
        clear2 = '/html/body/div/p[2]'
        clear_line2 = browser.find_element_by_xpath(clear2)
        ActionChains(browser).move_to_element(clear_line2).click(clear_line2).perform()
        Wait(2)
        clear_line2.clear()
        Wait(2)
        clear3 = '/html/body/div/p[1]'
        clear_line3 = browser.find_element_by_xpath(clear3)
        ActionChains(browser).move_to_element(clear_line3).click(clear_line3).send_keys(" ").perform()
        audit_log("REQ/DCA - Automagica - Inpatient - Removed terms and conditions for Credit Note.", "Completed...", reqdca_base)
      else:
        audit_log("REQ/DCA - Automagica - Inpatient - Terms and conditions were not removed.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Remove terms and conditions for Credit Note")
      logging("REQDCA - Automagica - Inpatient - Remove terms and conditions for Credit Note", error, reqdca_base)
    pass

  time.sleep(5)

  # Indicate Reason
  try:
    browser.find_element_by_xpath('/html/body/div/table[2]/tbody/tr[2]/td[1]/span').clear()
    element_reason = '/html/body/div/table[2]/tbody/tr[2]/td[1]/br[2]'
    myreason = wait_to_click(browser, element_reason)
    ActionChains(browser).move_to_element(myreason).click(myreason).send_keys("%s" % str(remarks)).perform()
    audit_log("REQ/DCA - Automagica - Inpatient - Indicate Reason.", "Completed...", reqdca_base)
  except Exception:
    try:
      myreason = wait_to_click(browser, element_reason)
      ActionChains(browser).move_to_element(myreason).click(myreason).send_keys("%s" % str(remarks)).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Indicate Reason.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Indicate Reason")
      logging("REQDCA - Automagica - Inpatient - Indicate Reason", error, reqdca_base)
    pass

  time.sleep(5)
  ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()

  # Inserting the Total Amount
  try:
    element_total = '/html/body/div/table[2]/tbody/tr[3]/td[2]'
    mytotal = wait_to_click(browser, element_total)
    ActionChains(browser).move_to_element(mytotal).click(mytotal).perform()
    time.sleep(3)
    if int(float(amount)) < int(0):
      print("Negative Amount")
      mytotal.clear()
      mytotal.click()
      mytotal.click()
      time.sleep(3)
      ActionChains(browser).move_to_element(mytotal).click(mytotal).send_keys("- RM %s " % str(amount).replace('-', '')).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Inserting the Total Amount [ - RM %s ]." % str(amount).replace('-', ''), "Completed...", reqdca_base)
    elif int(float(amount)) > int(0):
      print("Positive Amount")
      mytotal.clear()
      mytotal.click()
      mytotal.click()
      time.sleep(3)
      ActionChains(browser).move_to_element(mytotal).click(mytotal).send_keys("RM %s " % amount).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Inserting the Total Amount [ RM %s ]." % amount, "Completed...", reqdca_base)
  except Exception:
    try:
      element_total = '/html/body/div/table[2]/tbody/tr[3]/td[2]'
      mytotal = wait_to_click(browser, element_total)
      ActionChains(browser).move_to_element(mytotal).click(mytotal).perform()
      time.sleep(3)
      if int(float(amount)) < int(0):
        print("Negative Amount")
        mytotal.clear()
        mytotal.click()
        time.sleep(3)
        ActionChains(browser).move_to_element(mytotal).click(mytotal).send_keys("- RM %s " % str(amount).replace('-', '')).perform()
        audit_log("REQ/DCA - Automagica - Inpatient - Inserting the Total Amount [ - RM %s ]." % str(amount).replace('-', ''), "Completed...", reqdca_base)
      else:
        print("Positive Amount")
        mytotal.clear()
        mytotal.click()
        time.sleep(3)
        ActionChains(browser).move_to_element(mytotal).click(mytotal).send_keys("RM %s " % amount).perform()
        audit_log("REQ/DCA - Automagica - Inpatient - Inserting the Total Amount [ RM %s ]." % amount, "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Inserting Total Amount")
      logging("REQDCA - Automagica - Inpatient - Inserting Total Amount", error, reqdca_base)
    pass

  time.sleep(5)
  ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
  ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
  time.sleep(5)

  # Save and Send
  try:
    # Move to the main window
    browser.switch_to.default_content()
    element_ss = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/div[2]/div[2]/div/div/table/tbody/tr[2]/td/button'
    but_loc = wait_to_click(browser, element_ss)
    but_loc.click()
    audit_log("REQ/DCA - Automagica - Inpatient - Successfull document for Case ID [ %s ] created in MARC." % caseid, "Completed...", reqdca_base)
  except Exception:
    try:
      browser.switch_to.default_content()
      element_ss = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/div[2]/div[2]/div/div/table/tbody/tr[2]/td/button'
      but_loc = wait_to_click(browser, element_ss)
      ActionChains(browser).move_to_element(but_loc).click(but_loc).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Successfull document for Case ID [ %s ] created in MARC." % caseid, "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Save and Send")
      logging("REQDCA - Automagica - Inpatient - Save and Send", error, reqdca_base)
    pass

  time.sleep(5)

  # Close the window properly
  try:
    element_x = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/div[3]/div[1]/a'
    my_x = wait_to_click(browser, element_x)
    ActionChains(browser).move_to_element(my_x).click(my_x).perform()
    audit_log("REQ/DCA - Inpatient - Window is closed properly.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_x = '/html/body/div[1]/div[3]/form[1]/div[3]/div[2]/div/div/div[6]/div[3]/div[1]/a/span'
      my_x = wait_to_click(browser, element_x)
      ActionChains(browser).move_to_element(my_x).click(my_x).perform()
      audit_log("REQ/DCA - Inpatient - Window is closed properly.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Close X")
      logging("REQDCA - Automagica - Inpatient - Close the window properly", error, reqdca_base)
    pass

  time.sleep(5)

  # Save the record
  try:
    element_finalsave = '/html/body/div[1]/div[3]/form[1]/div[1]/div[1]/div[1]/a[5]'
    myfinalsave = wait_to_click(browser, element_finalsave)
    ActionChains(browser).move_to_element(myfinalsave).click(myfinalsave).perform()
    audit_log("REQ/DCA - Automagica - Inpatient - Successfull document for Case ID [ %s ] saved in MARC." % caseid, "Completed...", reqdca_base)
  except Exception:
    try:
      element_finalsave = '/html/body/div[1]/div[3]/form[1]/div[1]/div[1]/div[1]/a[5]/img'
      myfinalsave = wait_to_click(browser, element_finalsave)
      myfinalsave.click()
      audit_log("REQ/DCA - Automagica - Inpatient - Successfull document for Case ID [ %s ] saved in MARC." % caseid, "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Inpatient - Saved Record in MARC", error, reqdca_base)
    pass

  time.sleep(10)

  # Unlock the Case ID properly
  try:
    element_unlock = '/html/body/div[1]/div[3]/form[1]/div[1]/div[1]/div[2]/a[3]'
    my_unlock = wait_to_click(browser, element_unlock)
    my_unlock.click()
    audit_log("REQ/DCA - Automagica - Inpatient - Case ID [ %s ] has been unlocked." % caseid, "Completed...", reqdca_base)
  except Exception:
    try:
      element_unlock = '/html/body/div[1]/div[3]/form[1]/div[1]/div[1]/div[2]/a[3]'
      my_unlock = wait_to_click(browser, element_unlock)
      ActionChains(browser).move_to_element(my_unlock).click(my_unlock).perform()
      audit_log("REQ/DCA - Automagica - Inpatient - Case ID [ %s ] has been unlocked." % caseid, "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Unlock Case ID")
      logging("REQDCA - Automagica - Inpatient - Unlock the Case ID properly", error, reqdca_base)
    pass

  time.sleep(10)

def create_doc_outpatient(browser, attention, amount, remarks, HNS, doc_template, insurer, subcaseid, address, cln, current_date, current_time, reqdca_base):

  # Edit document of client
  try:
    element_edit = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/a[2]/img'
    myedit = wait_to_click(browser, element_edit)
    myedit.click()
  except Exception:
    try:
      element_edit = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/a[2]'
      myedit = wait_to_click(browser, element_edit)
      ActionChains(browser).move_to_element(myedit).click(myedit).perform()
    except Exception as error:
      logging("REQDCA - Automagica - Outpatient - Edit document of client", error, reqdca_base)
    pass

  time.sleep(30)

  try:
    element_doc = '/html/body/div[1]/div[3]/form/div[4]/div[2]/div/ul/li[5]/a'
    mydoc = wait_to_click(browser, element_doc)
    mydoc.click()
    audit_log("REQ/DCA - Automagica - Outpatient - Edit document of client.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_doc = '/html/body/div[1]/div[3]/form/div[4]/div[2]/div/ul/li[5]'
      mydoc = wait_to_click(browser, element_doc)
      ActionChains(browser).move_to_element(mydoc).click(mydoc).perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Edit document of client.", "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Outpatient - Edit document of client", error, reqdca_base)
    pass

  time.sleep(5)
    
  # Select option in dropdown menu
  try:
    Select(browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[4]/div[2]/div/div/div[5]/table[1]/tbody/tr[2]/td[2]/select')).select_by_visible_text(doc_template)
    Wait(2)
    Select(browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[4]/div[2]/div/div/div[5]/table[1]/tbody/tr[3]/td[4]/select')).select_by_visible_text(insurer)
    Wait(2)
    browser.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[4]/div[2]/div/div/div[5]/table[1]/tbody/tr[3]/td[2]/input').send_keys(attention)
    audit_log("REQ/DCA - Automagica - Outpatient - Select option in dropdown menu.", "Completed...", reqdca_base)
  except Exception as error:
    logging("REQDCA - Automagica - Outpatient - Select option in dropdown menu", error, reqdca_base)
    pass

  time.sleep(5)

  # Update LG amount
  try:
    element_lg = '/html/body/div[1]/div[3]/form/div[4]/div[2]/div/div/div[5]/table[1]/tbody/tr[2]/td[4]/input'
    mylgamt = wait_to_click(browser, element_lg)
    mylgamt.clear()
    ActionChains(browser).move_to_element(mylgamt).click(mylgamt).send_keys("%s" % amount).perform()
    audit_log("REQ/DCA - Automagica - Outpatient - Update LG amount.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_lg = '/html/body/div[1]/div[3]/form/div[4]/div[2]/div/div/div[5]/table[1]/tbody/tr[2]/td[4]/input'
      mylgamt = wait_to_click(browser, element_lg)
      mylgamt.clear()
      ActionChains(browser).move_to_element(mylgamt).click(mylgamt).send_keys("%s" % amount).perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Update LG amount.", "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Outpatient - Update LG amount", error, reqdca_base)
    pass

  time.sleep(5)

  # Select "Send as"
  try:
      element_save = '/html/body/div[1]/div[3]/form/div[4]/div[2]/div/div/div[5]/table[1]/tbody/tr[6]/td[2]/table/tbody/tr/td[1]/input'
      mysave = wait_to_click(browser, element_save)
      mysave.click()
      audit_log("REQ/DCA - Automagica - Outpatient - Select Send As.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_save = '/html/body/div[1]/div[3]/form/div[4]/div[2]/div/div/div[5]/table[1]/tbody/tr[6]/td[2]/table/tbody/tr/td[1]/input'
      mysave = wait_to_click(browser, element_save)
      ActionChains(browser).move_to_element(mysave).click(mysave).perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Select Send As.", "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Outpatient - Select Send As", error, reqdca_base)
    pass

  time.sleep(5)
  ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()
  time.sleep(5)

  # Create Document
  try:
    element_createdoc = '/html/body/div[1]/div[3]/form/div[4]/div[2]/div/div/div[5]/button'
    mycreatedoc = wait_to_click(browser, element_createdoc)
    mycreatedoc.click()
    audit_log("REQ/DCA - Automagica - Outpatient - Create Document.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_createdoc = '/html/body/div[1]/div[3]/form/div[4]/div[2]/div/div/div[5]/button'
      mycreatedoc = wait_to_click(browser, element_createdoc)
      ActionChains(browser).move_to_element(mycreatedoc).click(mycreatedoc).perform()
    except Exception as error:
      logging("REQDCA - Automagica - Outpatient - Create Document", error, reqdca_base)
    pass

  time.sleep(60)

  # Set the browser to popup window
  try:
    iframes = browser.find_elements_by_xpath("//iframe")
    browser.switch_to.frame(iframes[0])
    audit_log("REQ/DCA - Automagica - Outpatient - Set the browser to popup window.", "Completed...", reqdca_base)
  except Exception:
    try:
      iframes = browser.find_elements_by_xpath("//iframe")
      browser.switch_to.frame(iframes[0])
      audit_log("REQ/DCA - Automagica - Outpatient - Set the browser to popup window.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Set the browser to popup window")
      logging("REQDCA - Automagica - Outpatient - Set the browser to popup window", error, reqdca_base)
    pass

  time.sleep(5)

  # Edit 'bill to' field
  try:
    element_bill = '/html/body/div/h3[2]'
    mybill = wait_to_click(browser, element_bill)
    mybill.click()
    ActionChains(browser).move_to_element(mybill).click(mybill).send_keys("%s" % address).perform()
    audit_log("REQ/DCA - Automagica - Outpatient - Edit 'bill to' field.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_bill = '/html/body/div/h3[2]'
      mybill = wait_to_click(browser, element_bill)
      mybill.click()
      ActionChains(browser).move_to_element(mybill).click(mybill).send_keys("%s" % address).perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Set the Bill To.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Edit 'bill to' field")
      logging("REQDCA - Automagica - Outpatient - Set the Bill To", error, reqdca_base)
    pass

  time.sleep(5)

  # Inserting HNS running number
  try:
    element_hns = '/html/body/div/table[1]/tbody/tr[2]/td[3]'
    myhns = wait_to_click(browser, element_hns)
    ActionChains(browser).move_to_element(myhns).click(myhns).send_keys("%s" % str(HNS)).perform()
    audit_log("REQ/DCA - Automagica - Outpatient - Inserting HNS running number.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_hns = '/html/body/div/table[1]/tbody/tr[2]/td[3]'
      myhns = wait_to_click(browser, element_hns)
      ActionChains(browser).move_to_element(myhns).click(myhns).send_keys("%s" % str(HNS)).perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Inserting HNS running number.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Inserting HNS running number")
      logging("REQDCA - Automagica - Outpatient - Inserting HNS running number", error, reqdca_base)
    pass

  time.sleep(5)

  # Inserting Client Listing Number
  try:
    element_cln = '/html/body/div/table[1]/tbody/tr[2]/td[4]'
    mycln = wait_to_click(browser, element_cln)
    ActionChains(browser).move_to_element(mycln).click(mycln).send_keys("%s" % str(cln)).perform()
    audit_log("REQ/DCA - Automagica - Outpatient - Inserting Client Listing number.", "Completed...", reqdca_base)
  except Exception:
    try:
      element_cln = '/html/body/div/table[1]/tbody/tr[2]/td[4]'
      mycln = wait_to_click(browser, element_cln)
      ActionChains(browser).move_to_element(mycln).click(mycln).send_keys("%s" % str(cln)).perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Inserting Client Listing Number.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Inserting Client Listing number")
      logging("REQDCA - Automagica - Outpatient - Inserting Client Listing Number", error, reqdca_base)
    pass

  time.sleep(5)

  # Remove terms and conditions for Credit Note
  try:
    if int(float(amount)) < int(0):
      clear1 = '/html/body/div/p[1]'
      clear_line1 = browser.find_element_by_xpath(clear1)
      clear_line1.click()
      Wait(2)
      clear_line1.clear()
      Wait(2)
      clear2 = '/html/body/div/p[2]'
      clear_line2 = browser.find_element_by_xpath(clear2)
      ActionChains(browser).move_to_element(clear_line2).click(clear_line2).perform()
      Wait(2)
      clear_line2.clear()
      Wait(2)
      clear3 = '/html/body/div/p[1]'
      clear_line3 = browser.find_element_by_xpath(clear3)
      ActionChains(browser).move_to_element(clear_line3).click(clear_line3).send_keys(" ").perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Removed terms and conditions for Credit Note.", "Completed...", reqdca_base)
    else:
      audit_log("REQ/DCA - Automagica - Outpatient - Terms and conditions were not removed.", "Completed...", reqdca_base)
  except Exception:
    try:
      if int(float(amount)) < int(0):
        clear1 = '/html/body/div/p[1]'
        clear_line1 = browser.find_element_by_xpath(clear1)
        ActionChains(browser).move_to_element(clear_line1).click(clear_line1).perform()
        Wait(2)
        clear_line1.clear()
        Wait(2)
        clear2 = '/html/body/div/p[2]'
        clear_line2 = browser.find_element_by_xpath(clear2)
        ActionChains(browser).move_to_element(clear_line2).click(clear_line2).perform()
        Wait(2)
        clear_line2.clear()
        Wait(2)
        clear3 = '/html/body/div/p[1]'
        clear_line3 = browser.find_element_by_xpath(clear3)
        ActionChains(browser).move_to_element(clear_line3).click(clear_line3).send_keys(" ").perform()
        audit_log("REQ/DCA - Automagica - Outpatient - Removed terms and conditions for Credit Note.", "Completed...", reqdca_base)
      else:
        audit_log("REQ/DCA - Automagica - Outpatient - Terms and conditions were not removed.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Remove terms and conditions for Credit Note")
      logging("REQDCA - Automagica - Outpatient - Remove terms and conditions for Credit Note", error, reqdca_base)
    pass

  time.sleep(5)

  # Indicate Reason
  try:
    browser.find_element_by_xpath('/html/body/div/table[2]/tbody/tr[2]/td[1]/span').clear()
    element_reason = '/html/body/div/table[2]/tbody/tr[2]/td[1]/br[2]'
    myreason = browser.find_element_by_xpath(element_reason)
    ActionChains(browser).move_to_element(myreason).click(myreason).send_keys("%s" % str(remarks)).perform()
    audit_log("REQ/DCA - Automagica - Outpatient - Indicate Reason.", "Completed...", reqdca_base)
  except Exception:
    try:
      myreason = browser.find_element_by_xpath(element_reason)
      ActionChains(browser).move_to_element(myreason).click(myreason).send_keys("%s" % str(remarks)).perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Indicate Reason.", "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Indicate Reason")
      logging("REQDCA - Automagica - Outpatient - Indicate Reason", error, reqdca_base)
    pass

  time.sleep(5)

  ActionChains(browser).send_keys(Keys.PAGE_DOWN).perform()

  # Inserting the Total Amount
  try:
    element_total = '/html/body/div/table[2]/tbody/tr[3]/td[2]'
    mytotal = wait_to_click(browser, element_total)
    ActionChains(browser).move_to_element(mytotal).click(mytotal).perform()
    time.sleep(3)
    if int(float(amount)) < int(0):
      print("Negative Amount")
      mytotal.clear()
      mytotal.click()
      mytotal.click()
      time.sleep(3)
      ActionChains(browser).move_to_element(mytotal).click(mytotal).send_keys("- RM %s " % str(amount).replace('-', '')).perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Inserting the Total Amount [ - RM %s ]." % str(amount).replace('-', ''), "Completed...", reqdca_base)
    elif int(float(amount)) > int(0):
      print("Positive Amount")
      mytotal.clear()
      mytotal.click()
      mytotal.click()
      time.sleep(3)
      ActionChains(browser).move_to_element(mytotal).click(mytotal).send_keys("RM %s " % amount).perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Inserting the Total Amount [ RM %s ]." % amount, "Completed...", reqdca_base)
  except Exception:
    try:
      element_total = '/html/body/div/table[2]/tbody/tr[3]/td[2]'
      mytotal = wait_to_click(browser, element_total)
      ActionChains(browser).move_to_element(mytotal).click(mytotal).perform()
      time.sleep(3)
      if int(float(amount)) < int(0):
        print("Negative Amount")
        mytotal.clear()
        mytotal.click()
        time.sleep(3)
        ActionChains(browser).move_to_element(mytotal).click(mytotal).send_keys("- RM %s " % str(amount).replace('-', '')).perform()
        audit_log("REQ/DCA - Automagica - Outpatient - Inserting the Total Amount [ - RM %s ]." % str(amount).replace('-', ''), "Completed...", reqdca_base)
      else:
        print("Positive Amount")
        mytotal.clear()
        mytotal.click()
        time.sleep(3)
        ActionChains(browser).move_to_element(mytotal).click(mytotal).send_keys("RM %s " % amount).perform()
        audit_log("REQ/DCA - Automagica - Outpatient - Inserting the Total Amount [ RM %s ]." % amount, "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Inserting Total Amount")
      logging("REQDCA - Automagica - Outpatient - Inserting Total Amount", error, reqdca_base)
    pass

  time.sleep(5)
  ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
  ActionChains(browser).send_keys(Keys.PAGE_UP).perform()
  time.sleep(5)

  # Save and Send
  try:
    browser.switch_to.default_content()
    element_ss = '/html/body/div[1]/div[3]/form/div[4]/div[2]/div/div/div[5]/div[2]/div[2]/div/div/button'
    but_loc = wait_to_click(browser, element_ss)
    but_loc.click()
    audit_log("REQ/DCA - Automagica - Outpatient - Successfull document for Case ID [ %s ] created in MARC." % subcaseid, "Completed...", reqdca_base)
  except Exception:
    try:
      browser.switch_to.default_content()
      element_ss = '/html/body/div[1]/div[3]/form/div[4]/div[2]/div/div/div[5]/div[2]/div[2]/div/div/button'
      but_loc = wait_to_click(browser, element_ss)
      ActionChains(browser).move_to_element(but_loc).click(but_loc).perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Successfull document for Case ID [ %s ] created in MARC." % subcaseid, "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Outpatient - Save and Send", error, reqdca_base)
    pass

  time.sleep(5)

  # Save the record
  try:
    element_finalsave = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/a[2]/img'
    myfinalsave = wait_to_click(browser, element_finalsave)
    ActionChains(browser).move_to_element(myfinalsave).click(myfinalsave).perform()
    audit_log("REQ/DCA - Automagica - Outpatient - Successfull document for Case ID [ %s ] saved in MARC." % subcaseid, "Completed...", reqdca_base)
  except Exception:
    try:
      element_finalsave = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/a[2]'
      myfinalsave = wait_to_click(browser, element_finalsave)
      myfinalsave.click()
      audit_log("REQ/DCA - Automagica - Outpatient - Successfull document for Case ID [ %s ] saved in MARC." % subcaseid, "Completed...", reqdca_base)
    except Exception as error:
      logging("REQDCA - Automagica - Outpatient - Save Record in MARC", error, reqdca_base)

  time.sleep(10)

  # Unlock the Case ID properly
  try:
    element_unlock = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/a[3]'
    my_unlock = wait_to_click(browser, element_unlock)
    my_unlock.click()
    audit_log("REQ/DCA - Automagica - Outpatient - Case ID [ %s ] has been unlocked." % subcaseid, "Completed...", reqdca_base)
  except Exception:
    try:
      element_unlock = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/a[3]/img'
      my_unlock = wait_to_click(browser, element_unlock)
      ActionChains(browser).move_to_element(my_unlock).click(my_unlock).perform()
      audit_log("REQ/DCA - Automagica - Outpatient - Case ID [ %s ] has been unlocked." % subcaseid, "Completed...", reqdca_base)
    except Exception as error:
      print("- Error at Automagica section Unlock Case ID")
      logging("REQDCA - Automagica - Outpatient - Unlock the Case ID properly", error, reqdca_base)
    pass

  time.sleep(5)
