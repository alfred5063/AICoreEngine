#!/usr/bin/python
# FINAL SCRIPT updated as of 06st July 2020
# Workflow - Finance Credit Control

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
def login_Marc(base):
  try:

    global browser

    profile = webdriver.FirefoxProfile()

    # Check if both username and password are not null
    if base.username != '':
      browser = webdriver.Firefox(firefox_profile = profile, executable_path = dti_path+r'/selenium_interface/geckodriver.exe')
      browser.maximize_window()
      #browser.get('https://marcmy.asia-assistance.com/marcmy/')
      browser.get('http://192.168.2.173:8080/marcmy/')
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').clear()
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').send_keys(base.username)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').clear()
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').send_keys(base.password)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/button').click()
      audit_log("Finance Credit Control - Automagica - Accessing MARC using [ %s ]." % base.username, "Completed...", base)
    else:
      config = read_db_config(dti_path+r'\config.ini', 'marc-creditcontrol')
      dti_user = config['user']
      dti_pass = config['password']

      browser = webdriver.Firefox(firefox_profile = profile, executable_path = dti_path+r'/selenium_interface/geckodriver.exe')
      browser.maximize_window()
      #browser.get('https://marcmy.asia-assistance.com/marcmy/')
      browser.get('http://192.168.2.173:8080/marcmy/')
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').clear()
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[2]').send_keys(dti_user)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').clear()
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/input[3]').send_keys(dti_pass)
      Wait(2)
      browser.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/form/button').click()
      audit_log("Finance Credit Control - Automagica - Accessing MARC using [ %s ]." % dti_user, "Completed...", base)
    return browser
  except Exception as error:
    audit_log("Finance Credit Control - Web browser failed. MARC database cannot be accessed.", "Completed...", base)
    logging("Finance Credit Control - Automagica - Failed access to web browser.", error, base)

def navigate_page(browser, *args):

    #navigate to any page, refer to URL for parameter
    #example : http://192.168.2.173:8080/marcmy/pages/dashboard/dashboard_inpatient.xhtml
    #navigate_page(browser, "dashboard", "dashboard_inpatient")

    #page = 'https://marcmy.asia-assistance.com/marcmy/pages'
    page = 'http://192.168.2.173:8080/marcmy/pages'
    for arg in args:
        page = page + '/' + arg
    page = page + '.xhtml'
    return browser.get(page)

def close_Marc(browser):

  browser.close()

def get_post_address(browser):

  # Get IMPPID
  element_imppid = '/html/body/div[1]/div[3]/form/div[4]/div[1]/div/div/div[1]/table/tbody/tr[1]/td[2]/table/tbody/tr[1]/td[2]/input'
  imppid = browser.find_element_by_xpath(element_imppid).get_attribute("value")
  time.sleep(5)

  element_imppid = '/html/body/div[1]/div[3]/form/div[4]/div[1]/div/div/div[1]/table/tbody/tr[1]/td[2]/table/tbody/tr[1]/td[2]/input'
  imppid = browser.find_element_by_xpath(element_imppid).get_attribute("value")
  time.sleep(5)

  if imppid != '':
    browser.get('http://192.168.2.173:8080/marcmy/pages/setup/inptmember/memberpolicy_search.xhtml')
    time.sleep(5)

    element_searchimppid = '/html/body/div[1]/div[3]/form/table[1]/tbody/tr[1]/td[2]/input'
    searchimppid = browser.find_element_by_xpath(element_searchimppid)
    ActionChains(browser).move_to_element(searchimppid).click(searchimppid).send_keys("%s" % imppid).perform()
    time.sleep(10)

    element_button = '/html/body/div[1]/div[3]/form/table[1]/tbody/tr[8]/td[2]/button/span[2]'
    clickbutton = browser.find_element_by_xpath(element_button)
    ActionChains(browser).move_to_element(clickbutton).click(clickbutton).perform()
    time.sleep(10)

    element_record = '/html/body/div[1]/div[3]/form/div/div/div/div[3]/table/tbody/tr/td[1]/button'
    view_record = browser.find_element_by_xpath(element_record)
    ActionChains(browser).move_to_element(view_record).click(view_record).perform()
    time.sleep(20)

    try:
      element_add1 = '/html/body/div[1]/div[3]/form/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr[1]/td[2]'
      copy_add1 = browser.find_element_by_xpath(element_add1).get_attribute("textContent")
      element_add2 = '/html/body/div[1]/div[3]/form/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr[2]/td[2]'
      copy_add2 = browser.find_element_by_xpath(element_add2).get_attribute("textContent")
      element_add3 = '/html/body/div[1]/div[3]/form/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr[3]/td[2]'
      copy_add3 = browser.find_element_by_xpath(element_add3).get_attribute("textContent")

      if copy_add1 == '':
        patient_address = "Member's mailing address is not recorded in MARC. Please add manually."
      elif copy_add1 != '' and copy_add2 == '':
        patient_address = str(copy_add1) + " (Member's mailing address is not complete in MARC. Please add manually)."
      elif copy_add1 != '' and element_add2 != '' and copy_add3 != '' or copy_add3 == '':
        patient_address = str(copy_add1) + ', ' + str(copy_add2) + ', ' + str(copy_add3)

      if patient_address != '':
        audit_log("Finance Credit Control - Member's address is [ %s ]." % patient_address, "Completed...", creditcontrol_base)
      else:
        patient_address = str(patient_address)

    except Exception:
      time.sleep(5)
      element_add1 = '/html/body/div[1]/div[3]/form/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr[1]/td[2]'
      copy_add1 = browser.find_element_by_xpath(element_add1).get_attribute("textContent")
      element_add2 = '/html/body/div[1]/div[3]/form/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr[2]/td[2]'
      copy_add2 = browser.find_element_by_xpath(element_add2).get_attribute("textContent")
      element_add3 = '/html/body/div[1]/div[3]/form/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr[3]/td[2]'
      copy_add3 = browser.find_element_by_xpath(element_add3).get_attribute("textContent")

      if copy_add1 == '':
        patient_address = "Member's mailing address is not recorded in MARC. Please add manually."
      elif copy_add1 != '' and copy_add2 == '':
        patient_address = str(copy_add1) + " (Member's mailing address is not complete in MARC. Please add manually)."
      elif copy_add1 != '' and element_add2 != '' and copy_add3 != '' or copy_add3 == '':
        patient_address = str(copy_add1) + ', ' + str(copy_add2) + ', ' + str(copy_add3)

    element_logout = '/html/body/div[1]/div[1]/div[2]/form/a[2]'
    logout = browser.find_element_by_xpath(element_logout)
    ActionChains(browser).move_to_element(logout).click(logout).perform()
    time.sleep(5)

  else:
     patient_address = "Unable to extract Member's mailing address from MARC. Please add manually."

  return patient_address

def search_post_member(browser, file_no, creditcontrol_base):

  # Send Case ID
  element_sendid = '/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[1]/td[4]/input'
  sendid = browser.find_element_by_xpath(element_sendid)
  ActionChains(browser).move_to_element(sendid).click(sendid).send_keys("%s" % file_no).perform()
  time.sleep(5)

  # Click Button Search
  element_search = '/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[7]/td[2]/button[1]'
  sendid = browser.find_element_by_xpath(element_search)
  ActionChains(browser).move_to_element(sendid).click(sendid).perform()
  time.sleep(5)

  # Click Sub Case ID
  element_subcase = '/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[1]/a'
  subcase = browser.find_element_by_xpath(element_subcase)
  ActionChains(browser).move_to_element(subcase).click(subcase).perform()
  time.sleep(5)

  # Get IC Number
  try:
    element_ic = '/html/body/div[1]/div[3]/form[1]/div[3]/div[1]/div/div/div[1]/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/input'
    icno = browser.find_element_by_xpath(element_ic).get_attribute("value")
    time.sleep(5)
    if icno != '':
      audit_log("Finance Credit Control - Member's IC Number is [ %s ]." % icno, "Completed...", creditcontrol_base)
      send(creditcontrol_base, 'choonlam.chin@asia-assistance.com', "IC number is empty.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(creditcontrol_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
    else:
      icno = str(icno)
  except Exception:
    time.sleep(10)
    element_ic = '/html/body/div[1]/div[3]/form[1]/div[3]/div[1]/div/div/div[1]/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/input'
    icno = browser.find_element_by_xpath(element_ic).get_attribute("value")
    time.sleep(5)
    if icno == '':
      send(creditcontrol_base, 'choonlam.chin@asia-assistance.com', "IC number is empty.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(creditcontrol_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
    else:
      icno = str(icno)

  return icno

def search_adm_member(browser, file_no, creditcontrol_base):

  # Send Case ID
  element_sendid = '/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[1]/td[4]/input'
  sendid = browser.find_element_by_xpath(element_sendid)
  ActionChains(browser).move_to_element(sendid).click(sendid).send_keys("%s" % file_no).perform()
  time.sleep(5)

  # Click Button Search
  element_search = '/html/body/div[1]/div[3]/form/div[2]/table[1]/tbody/tr[11]/td[2]/button[1]'
  sendid = browser.find_element_by_xpath(element_search)
  ActionChains(browser).move_to_element(sendid).click(sendid).perform()
  time.sleep(5)

  # Get IC Number
  try:
    element_ic = '/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[5]'
    icno = browser.find_element_by_xpath(element_ic).text
    time.sleep(5)
    if icno != '':
      audit_log("Finance Credit Control - Member's IC Number is [ %s ]." % icno, "Completed...", creditcontrol_base)
      send(creditcontrol_base, 'choonlam.chin@asia-assistance.com', "IC number is empty.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(creditcontrol_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
    else:
      icno = str(icno)
  except Exception:
    time.sleep(10)
    element_ic = '/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[5]'
    icno = browser.find_element_by_xpath(element_ic).text
    time.sleep(5)
    if icno == '':
      send(creditcontrol_base, 'choonlam.chin@asia-assistance.com', "IC number is empty.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(creditcontrol_base.guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
    else:
      icno = str(icno)

  return icno

def get_adm_address(browser):

  # Click Name
  element_name = '/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[4]/a'
  caseid = browser.find_element_by_xpath(element_name)
  ActionChains(browser).move_to_element(caseid).click(caseid).perform()
  time.sleep(5)

  element_imppid = '/html/body/div[1]/div[3]/form[1]/div[3]/div[1]/div/div/div[1]/table/tbody/tr[1]/td[2]/table/tbody/tr[1]/td[2]/input'
  imppid = browser.find_element_by_xpath(element_imppid).get_attribute("value")
  imppid = str(imppid)
  #imppid = browser.find_element_by_xpath(element_imppid).get_attribute("attr_name")
  print(imppid)
  #audit_log("Finance Credit Control - Imppid is [ %s ]." % str(imppid), "Completed...", creditcontrol_base)
  time.sleep(5)

  if imppid != '':
    print(imppid)
    browser.get('http://192.168.2.173:8080/marcmy/pages/setup/inptmember/memberpolicy_search.xhtml')
    time.sleep(5)

    element_searchimppid = '/html/body/div[1]/div[3]/form/table[1]/tbody/tr[1]/td[2]/input'
    searchimppid = browser.find_element_by_xpath(element_searchimppid)
    ActionChains(browser).move_to_element(searchimppid).click(searchimppid).send_keys("%s" % imppid).perform()
    time.sleep(10)

    element_button = '/html/body/div[1]/div[3]/form/table[1]/tbody/tr[8]/td[2]/button/span[2]'
    clickbutton = browser.find_element_by_xpath(element_button)
    ActionChains(browser).move_to_element(clickbutton).click(clickbutton).perform()
    time.sleep(10)

    element_record = '/html/body/div[1]/div[3]/form/div/div/div/div[3]/table/tbody/tr/td[1]/button'
    view_record = browser.find_element_by_xpath(element_record)
    ActionChains(browser).move_to_element(view_record).click(view_record).perform()
    time.sleep(20)

    element_add1 = '/html/body/div[1]/div[3]/form/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr[1]/td[2]'
    copy_add1 = browser.find_element_by_xpath(element_add1).get_attribute("textContent")
    element_add2 = '/html/body/div[1]/div[3]/form/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr[2]/td[2]'
    copy_add2 = browser.find_element_by_xpath(element_add2).get_attribute("textContent")
    element_add3 = '/html/body/div[1]/div[3]/form/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr[3]/td[2]'
    copy_add3 = browser.find_element_by_xpath(element_add3).get_attribute("textContent")

    if copy_add1 == '':
      patient_address = "Member's mailing address is not recorded in MARC. Please add manually."
    elif copy_add1 != '' and copy_add2 == '':
      patient_address = str(copy_add1) + " (Member's mailing address is not complete in MARC. Please add manually)."
    elif copy_add1 != '' and element_add2 != '' and copy_add3 != '' or copy_add3 == '':
      patient_address = str(copy_add1) + ', ' + str(copy_add2) + ', ' + str(copy_add3)

    element_logout = '/html/body/div[1]/div[1]/div[2]/form/a[2]'
    logout = browser.find_element_by_xpath(element_logout)
    ActionChains(browser).move_to_element(logout).click(logout).perform()
    time.sleep(5)

  else:
     patient_address = ""

  return patient_address

