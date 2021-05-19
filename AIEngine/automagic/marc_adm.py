#!/usr/bin/python
# FINAL SCRIPT updated as of 23rd June 2020
# Workflow - CBA Admission

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
from selenium.webdriver.support import expected_conditions as EC
from automagic.automagica_sql_setting import *
from utils.application import *
import os
from pathlib import Path
import pyautogui
from configparser import ConfigParser
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException,TimeoutException#SessionNotCreatedException
from connector.connector import MSSqlConnector,MySqlConnector
import json
from datetime import date,datetime,timedelta
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import traceback

current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)

app_overlay = get_application_url('marc_overlay')
overlay_id_list=[]
overlay_xpath_list=[]
try:
    counter=1
    overlay_id="o"
    while overlay_id!=None:
      overlay_id=get_xpath_by_key(app_overlay.application.appId, 'overlay_id_{}'.format(counter))
      if overlay_id !="None" and overlay_id !=None:
        overlay_id_list.append(overlay_id)
      counter+=1
except:
    pass
try:
    counter=1
    overlay_xpath="o"
    while overlay_xpath!=None:
      overlay_xpath=get_xpath_by_key(app_overlay.application.appId, 'overlay_{}'.format(counter))
      if overlay_xpath !="None" and overlay_xpath !=None:
        overlay_xpath_list.append(overlay_xpath)
      counter+=1
except:
    pass

def login_Marc(base, download_new = None):
    # Initiate browser
    """
    browser.download.folderList tells it not to use default Downloads directory
    browser.download.manager.showWhenStarting turns of showing download progress
    browser.download.dir sets the directory for downloads
    browser.helperApps.neverAsk.saveToDisk tells Firefox to automatically download the files of the selected mime-types
    """
    current_path = os.path.dirname(os.path.realpath(__file__))
    profile = webdriver.FirefoxProfile()
    if download_new is not None:
      profile.set_preference('browser.download.dir', download_new)
    profile.set_preference('browser.download.folderList', 2) # custom location
    profile.set_preference('browser.download.manager.showWhenStarting', False)
    profile.set_preference("browser.download.useDownloadDir", True)
    profile.set_preference("browser.download.manager.alertOnEXEOpen", False)
    profile.set_preference("browser.download.manager.closeWhenDone", False)
    profile.set_preference("browser.download.manager.focusWhenStarting", False)
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/octet-stream;charset=UTF-8;application/vnd.ms-excel')#vnd.ms-excel
    profile.set_preference("print.always_print_silent", True)
    profile.set_preference("browser.download.forbid_open_with", True)
    profile.set_preference("print.show_print_progress", False)
    profile.update_preferences()

    app_marc_login = get_application_url('marc_login')
    browser = webdriver.Firefox(firefox_profile = profile, executable_path = dti_path + r'/selenium_interface/geckodriver.exe')
    browser.get(app_marc_login.application.appUrl)

    try:
      if base.username != "" and base.password != "":
        print(len(base.password))
        browser.find_element_by_xpath(get_xpath_by_key(app_marc_login.application.appId, 'username')).send_keys(base.username)
        browser.find_element_by_xpath(get_xpath_by_key(app_marc_login.application.appId, 'password')).send_keys(base.password)
        browser.find_element_by_xpath(get_xpath_by_key(app_marc_login.application.appId, 'login_button')).click()
        wait_to_disappear(browser,get_xpath_by_key(app_marc_login.application.appId, 'login_button'))
      else:
        parser = ConfigParser()
        path = os.path.dirname(os.path.abspath(os.path.dirname(__file__)))
        filename=path+r"\config.ini"
        print(filename)
        parser.read(filename)
        if parser.has_section("marc2"):
          items = parser.items("marc2")
        browser.find_element_by_xpath(get_xpath_by_key(app_marc_login.application.appId, 'username')).send_keys(items[0][1])
        browser.find_element_by_xpath(get_xpath_by_key(app_marc_login.application.appId, 'password')).send_keys(items[1][1])
        browser.find_element_by_xpath(get_xpath_by_key(app_marc_login.application.appId, 'login_button')).click()
        wait_to_disappear(browser,get_xpath_by_key(app_marc_login.application.appId, 'login_button'))
        audit_log("Login in Marc", "Completed...", base)
    except NoSuchElementException:
      browser.get('http://192.168.2.173:8080/marcmy/')
    return browser

def go_Inpatient_Admission_View(browser):
  #go to admission view page without clicking
  app_inpatient = get_application_url('marc_search_inpatient')
  browser.get(app_inpatient.application.appUrl)
  

def go_Inpatient_Admission_Post_View(browser):
  #go to admission view page without clicking
  app_outpatient= get_application_url('marc_search_outpatient')
  browser.get(app_outpatient.application.appUrl)

  browser
  element = get_xpath_by_key(app_outpatient.application.appId, 'search_id')

def wait_to_load_long(browser, element):
  #wait for element to be show,suggested use in loading page for browser
  wait = WebDriverWait(browser, 20)
  wait.until(EC.visibility_of_element_located((By.XPATH, element)))

def short_wait_to_load(browser,element):
  #wait for element to be show,suggested for shorting and manipulating element
  wait=WebDriverWait(browser,3)
  wait.until(EC.visibility_of_element_located((By.XPATH, element)))

def wait_to_click(browser,element):
  #wait for element to be show,suggested use in JAVAscript loading page for browser
  wait=WebDriverWait(browser,20)
  wait.until(EC.element_to_be_clickable((By.XPATH, element)))


def key_InID (admNoList, browser):
    #insert case ID from list, search, and view patient,#no need wait
    app_outpatient = get_application_url('marc_search_outpatient')
    app_inpatient = get_application_url('marc_search_inpatient')
    wait_to_load_long(browser, get_xpath_by_key(app_outpatient.application.appId, 'search_id'))
    try:
      browser.find_element_by_xpath(get_xpath_by_key(app_outpatient.application.appId, 'search_id')).send_keys(admNoList)
    except:
      element=get_xpath_by_key(app_inpatient.application.appId, 'search_id')
      browser.find_element_by_xpath(element).send_keys(admNoList)

    try:
      browser.find_element_by_xpath(get_xpath_by_key(app_outpatient.application.appId, 'search_button')).click()
    except:
      try:
        app_inpatient = get_application_url('marc_search_inpatient')
        browser.find_element_by_xpath(get_xpath_by_key(app_inpatient.application.appId, 'search_button')).click()
      except:
        app_outpatient = get_application_url('marc_search_outpatient')
        browser.find_element_by_xpath(get_xpath_by_key(app_outpatient.application.appId, 'search_button')).click()

    
def click_PaitientName(browser):
  #Click paitent name to enter paitient profile ,require waiting
  #element change to 163 updated on 27/08/2019 change to 164 on 05/09
  #element = '/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[3]'
  element = "/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[1]/a"
  app_inpatient = get_application_url('marc_search_inpatient')
  try:
      short_wait_to_load(browser, element)
  except TimeoutException:
      element=get_xpath_by_key(app_inpatient.application.appId, 'member_entry')
  wait_to_load_long(browser,element)
  browser.find_element_by_xpath(element).click()


def click_PaitientName_post(browser):
  #Click paitent name to enter paitient profile ,require waiting
  app_outpatient = get_application_url('marc_search_outpatient')
  element = "/html/body/div[1]/div[3]/form/div[2]/div/div/div/div[3]/table/tbody/tr/td[1]/a"
  try:
      short_wait_to_load(browser,element)
  except TimeoutException:
      element=get_xpath_by_key(app_outpatient.application.appId, 'member_entry')
  wait_to_load_long(browser,element)
  browser.find_element_by_xpath(element).click()

def click_Ws(browser):
  #click Ws inside the paitient profile
  app_inpatient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpatient.application.appId, 'ws_tab')
  wait_to_click(browser,element)
  browser.find_element_by_xpath(element).click()

def click_CBA_Info(browser):
  #click Ws inside the paitient profile
  app_inpatient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpatient.application.appId, 'ws_cba_info')
  wait_to_click(browser,element)
  browser.find_element_by_xpath(element).click()
  
def click_Img(browser):
  #click Img inside the paitient profile
  app_inpatient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpatient.application.appId, 'img_tab')
  wait_to_click(browser,element)
  browser.find_element_by_xpath(element).click()

def check_Amount(number1,number2):
  #compare amount of GL and original resit ammount
  if number1==number2 or number2=="":
    return True
  else:
    return False

def check_exists_by_xpath(browser,xpath):
    try:
        browser.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def check_Status_post(browser):
  app_outpatient = get_application_url('marc_search_outpatient')
  return check_exists_by_xpath(browser,get_xpath_by_key(app_outpatient.application.appId, "status_ready_for_bord"))


def check_Status(browser,post=None):
  #check the status of application
  app_inpatient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpatient.application.appId, 'status')
  Status=browser.find_element_by_xpath(element).get_attribute("value")
  print(Status)
  if post=="POST":
    if Status=="CBA - Ready for Bdx":
      return True
    else:
      return False
  if Status=="CBA - Completed":
    return True
  elif Status== "TPA - Completed":
    return True
  elif Status== "Var - Completed":
    return True
  elif Status== "CBA - Ready for Bdx":
    return True
  else:
    return False

def get_Client_Name(browser):
  #get the client name , not billing name, require wait because come back from other page
  app_inpatient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpatient.application.appId, 'client')
  wait_to_load_long(browser,element)
  Client=browser.find_element_by_xpath(element).get_attribute("value")
  return Client

def get_Client_Name_post(browser):
  #get the client name , not billing name, require wait because come back from other page
  app_outpatient = get_application_url('marc_search_outpatient')
  element=get_xpath_by_key(app_outpatient.application.appId, 'client')
  wait_to_load_long(browser,element)
  Client=browser.find_element_by_xpath(element).get_attribute("value")
  return Client


def check_Doc(browser,id):
  #go to Document uploaded Iframe to check if there is any scanned file, and return back to paitient info
  click_Doc(browser)
  print("gg")
  try:
    app_inpatient = get_application_url('marc_search_inpatient')
    element=get_xpath_by_key(app_inpatient.application.appId, 'img_tab_iframe')
    iframe = browser.find_element_by_xpath(element)
    browser.switch_to.frame(iframe)
    if browser.find_element_by_class_name(get_xpath_by_key(app_inpatient.application.appId, 'img_tab_iframe_file_1_class'))  or browser.find_element_by_class_name(get_xpath_by_key(app_inpatient.application.appId, 'img_tab_iframe_file_2_class')):
      browser.switch_to.default_content()
      return True
    browser.switch_to.default_content()
    return False
  except:
    browser.get('https://doctrixsvr1.asia-assistance.com/doctrix/login.aspx?CaseID=369750&APICred=marc&NavTarget=API_AA_CUSTOMSEARCHVIEWER&IsUserExternal=0'.replace('369750',str(id)))
    app_inpatient = get_application_url('marc_search_inpatient')
    if browser.find_element_by_class_name(get_xpath_by_key(app_inpatient.application.appId, 'img_tab_iframe_file_1_class')) or browser.find_element_by_class_name(get_xpath_by_key(app_inpatient.application.appId, 'img_tab_iframe_file_2_class')):
      browser.get("http://192.168.2.173:8080/marcmy/pages/inpatient/cases/inpatient_cases_view.xhtml?faces-redirect=true")
      return True
    browser.get("http://192.168.2.173:8080/marcmy/pages/inpatient/cases/inpatient_cases_view.xhtml?faces-redirect=true")
    return False

def check_double_billing(browser):
# Logics
#Double Billing
#- check comp pay !=0
#- Bill to != Pol Owner

  client_value = get_Client_Name(browser)
  bill_to_value = get_Bill_to(browser)
  comp_pay = browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/form[1]/div[3]/div[2]/div[1]/div[1]/div[4]/table[1]/tbody[1]/tr[6]/td[6]/input[1]").get_attribute("value")
  b2bchecker = browser.find_element_by_xpath('//*[@id="cbBacktoBack"]').get_attribute("value")
  if client_value != bill_to_value and comp_pay != "" and float(comp_pay) > 0:
    # Double Billing
    return True
  else:
    return False

def check_double_billing_post(browser):
# Logics
#Double Billing
#- check comp pay !=0
#- Bill to != Pol Owner

  bill_to = get_Bill_to_post(browser)
  client_value = get_Client_Name_post(browser)
  comp_pay = get_comp_pay_post(browser)
  if client_value != bill_to and comp_pay != "" and float(comp_pay) > 0:
    return True
  else:
    return False

def get_Bill_to(browser):
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'bill_to')
  wait_to_click(browser,element)
  bill=browser.find_element_by_xpath(element).get_attribute("value")
  return bill

def get_Bill_to_post(browser):
  app_outpatient = get_application_url('marc_search_outpatient')
  element=get_xpath_by_key(app_outpatient.application.appId, 'bill_to')
  wait_to_click(browser,element)
  bill=browser.find_element_by_xpath(element).get_attribute("value")
  return bill

def get_comp_pay(browser):
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'comp_pay')
  #wait_to_click(browser,element)
  bill=browser.find_element_by_xpath(element).get_attribute("value")
  return bill

def get_comp_pay_post(browser):
  app_outpaitient = get_application_url('marc_search_outpatient')
  element=get_xpath_by_key(app_outpaitient.application.appId, 'company_pay')
  bill=browser.find_element_by_xpath(element).get_attribute("value")
  return bill

def click_Edit(browser):
  #click edit button at paitient page
  wait_for_overlag(browser)
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'edit_case')
  wait_to_click(browser,element)
  browser.find_element_by_xpath(element).click()

def click_Doc(browser):
  #click doc button at paitient page
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'doc_tab')
  wait_to_click(browser,element)
  browser.find_element_by_xpath(element).click()


def wait_to_disappear(browser,element):
  #wait for element to be show,suggested use in JAVAscript loading page for browser
  wait=WebDriverWait(browser,20)
  wait.until(EC.invisibility_of_element_located((By.XPATH, element)))


def wait_for_overlag(browser):
  print("start",datetime.now())
  wait_overlag=WebDriverWait(browser,200)
  for i in overlay_id_list:
    wait_overlag.until(EC.invisibility_of_element_located((By.ID,i)))
  for j in overlay_xpath_list:
    wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,j)))
  print("end",datetime.now())

#def wait_for_overlag(browser):
#  #wait for the webpage to load the overlay, overlay is the delay that make mouse cannot click
#  print("start",datetime.now())
#  
#  app_overlay = get_application_url('marc_overlay')
#  overlay_id="o"
#  overlay_xpath="o"
#  try:
#    counter=1
#    while overlay_id!=None:
#      overlay_id=get_xpath_by_key(app_overlay.application.appId, 'overlay_id_{}'.format(counter))
#      wait_overlag.until(EC.invisibility_of_element_located((By.ID,overlay_id)))
#      counter+=1
#  except:
#    pass

#  try:
#    counter=1
#    while overlay_xpath!=None:
#      overlay_xpath=get_xpath_by_key(app_overlay.application.appId, 'overlay_{}'.format(counter))
#      wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,overlay_xpath)))
#      counter+=1
#  except:
#    pass
#  print("end",datetime.now())
def find_client(browser,Client):
  """
  select client/insurer in Marc in bdx filling page
  """
  app_bord = get_application_url('marc_bord_creation')

  dropdown = Select(browser.find_element_by_id(get_xpath_by_key(app_bord.application.appId, 'client_id')))
  options = dropdown.options
  i=0
  for index in options:
    #need change to non linear search method
    i=i+1
    if Client in index.text:
      browser.find_element_by_xpath(get_xpath_by_key(app_bord.application.appId, 'client').format(i)).click()
      #ALERT might have 4 choice but same company
      break  
  
def select_First_Call(browser):
  #LONPAC,select first call, need wait due to javas , click to cause overlay
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'doc_select_first_call')
  wait_to_click(browser,element)
  browser.find_element_by_xpath(element).click()

def select_Insurer(browser):
  wait_for_overlag(browser)
  #LONPAC,select Lonpac, need to wait due to javas,click to cause overlay
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'doc_select_insurer')
  browser.find_element_by_xpath(element).click()

def click_None_Doc(browser):
  #LONPAC,select None,need to wait due to javas,click to cause overlay
  wait_for_overlag(browser)
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'doc_tab_none_id')
  browser.find_element_by_id(element).click()

def change_to_BdxReady(browser):
  #Change paitient status to CBA ready
  wait_for_overlag(browser)
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'status_select_ready_to_bord')
  browser.find_element_by_xpath(element).click()

def click_Save(browser):
  wait_for_overlag(browser)
  app_inpaitient = get_application_url('marc_search_inpatient')
  element_save=get_xpath_by_key(app_inpaitient.application.appId, 'save')
  box_id=get_xpath_by_key(app_inpaitient.application.appId, 'save_confirmation_box_id')
  element_yes_button=get_xpath_by_key(app_inpaitient.application.appId, 'save_confirmation_box_yes_button')
  #click Save button after edit inside the paitient profile
  try:
    wait_to_click(browser,element_save)
    browser.find_element_by_xpath(element_save).click()
  except:
    try:
      wait_save=WebDriverWait(browser,10)
      wait_save.until(EC.invisibility_of_element_located((By.ID,box_id)))
      browser.find_element_by_xpath(element_save).click()
    except:
      browser.find_element_by_xpath(element_save).click()
  try:
    wait_for_overlag(browser)
    wait_pop_out=WebDriverWait(browser,3)
    wait_pop_out.until(EC.visibility_of_element_located((By.XPATH,element_yes_button)))
    web_element=browser.find_element_by_xpath(element_yes_button)
    browser.find_element_by_xpath(element_yes_button).click()
    try:
      browser.find_element_by_xpath(element_yes_button).click
      web_element=browser.find_element_by_xpath(element_yes_button)
    except:
      pass
    ActionChains(browser).send_keys(Keys.ENTER).perform()
    ActionChains(browser).click(web_element).perform()
    Wait(8)
    click_unlock(browser)
  except Exception as e:
    Wait(8)
    click_unlock(browser)
  
def click_unlock(browser):
  wait_for_overlag(browser)
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'unlock')
  element2=get_xpath_by_key(app_inpaitient.application.appId, 'edit_case')
  try:
    wait_to_click(browser,element)
    browser.find_element_by_xpath(element).click()
    wait_to_click(browser,element2)
    wait_for_overlag(browser)
  except:
    pass

def go_Bordereaux_Add_New(browser):
  #go to Bordereaux_Add_New page without clicking
  app_bord_creation = get_application_url('marc_bord_creation')
  return browser.get(app_bord_creation.application.appUrl)

def get_ClientInsurance_Name(browser):
  #get the client name , not billing name, require wait because come back from other page
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'client')
  wait_to_load_long(browser,element)##specific to wait to load to paitient info,due to cannot seek different elemnet at same time
  Client=browser.find_element_by_xpath(element).get_attribute("value")
  return Client

def get_ClientInsurance_dbName(browser):
  #get the client name , not billing name, require wait because come back from other page
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'bill_to')
  wait_to_load_long(browser,element)##specific to wait to load to paitient info,due to cannot seek different elemnet at same time
  Client=browser.find_element_by_xpath(element).get_attribute("value")
  return Client

def bdx_Cashless_Admission(browser):
  #border creation, for selecting bdx type
  wait_for_overlag(browser)
  app_bord = get_application_url('marc_bord_creation')
  element=get_xpath_by_key(app_bord.application.appId, 'type_cashless')
  browser.find_element_by_xpath(element).click()

def bdx_Post_Cashless_Admission(browser):
  #border creation, for selecting bdx type
  wait_for_overlag(browser)
  app_bord = get_application_url('marc_bord_creation')
  element=get_xpath_by_key(app_bord.application.appId, 'type_cashless_post')
  browser.find_element_by_xpath(element).click()

def Check_B2B(browser,comp_pay,type):

# Logics
#Admission
#- check b2b column in MARC
#- check comp pay !=0

#Post
#- Bill to == Pol Owner
#- If B2B, check comp pay !=0
#- If not B2B, Ins Pay !=0 (Ws tab)

  status = False
  wait_for_overlag(browser)

  if type == 'POST':
    client_value = get_Client_Name_post(browser)
    bill_to_value = get_Bill_to_post(browser)
    comp_pay = get_comp_pay_post(browser)
  else:
    client_value = get_Client_Name(browser)
    bill_to_value = get_Bill_to(browser)
    comp_pay = get_comp_pay(browser)
  
  #b2bchecker = browser.find_element_by_xpath('//*[@id="cbBacktoBack"]').get_attribute("value")
  if str(client_value) == str(bill_to_value) and float(comp_pay) != '' and float(comp_pay) > int(0):
    status = True
  else:
    status = False
  return status


def bdx_Approved(browser):
  #border creation, for selecting claim type
  wait_for_overlag(browser)
  app_bord = get_application_url('marc_bord_creation')
  element=get_xpath_by_key(app_bord.application.appId, 'claim_type_approved')
  print(element)
  browser.find_element_by_xpath(element).click()

def bdx_Save(browser):
  #border creation, for selecting claim type
  wait_for_overlag(browser)
  app_inpatient = get_application_url('marc_bord_creation')
  element=get_xpath_by_key(app_inpatient.application.appId, 'save')
  browser.find_element_by_xpath(element).click()
  wait_for_overlag(browser)
  browser.find_element_by_xpath(element).click()
  wait_for_overlag(browser)

def bdx_Download(browser):
  wait_for_overlag(browser)
  app_bord = get_application_url('marc_bord_creation')
  #border creation, for selecting claim type
  try :
    element=get_xpath_by_key(app_bord.application.appId, 'download')
    browser.find_element_by_xpath(element).click()
    return True
  except NoSuchElementException:
    return False
  wait_for_overlag(browser)

def Download_WS(browser):
  #border creation, for selecting claim type
  wait_for_overlag(browser)
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'calc_first_ws')
  wait_to_click(browser,element)
  browser.find_element_by_xpath(element).click()

def Click_Calc(browser):
  wait_for_overlag(browser)
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'calc_tab')
  wait_to_click(browser,element)
  browser.find_element_by_xpath(element).click()
  
  
def get_Pol_Type(browser):
  #get policy type
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'policy_type')
  wait_to_load_long(browser,element)##specific to wait to load to paitient info,due to cannot seek different elemnet at same time
  Policy=browser.find_element_by_xpath(element).get_attribute("value")
  return Policy

def get_Pol_Type_post(browser):
  #get policy type
  app_outpaitient = get_application_url('marc_search_outpatient')
  element=get_xpath_by_key(app_outpaitient.application.appId, 'policy_type')
  wait_to_load_long(browser,element)##specific to wait to load to paitient info,due to cannot seek different elemnet at same time
  Policy=browser.find_element_by_xpath(element).get_attribute("value")
  return Policy

def bdx_Disb_No(browser,no):
  #border creation, for selecting claim type
  wait_for_overlag(browser)
  app_bord = get_application_url('marc_bord_creation')
  element=get_xpath_by_key(app_bord.application.appId, 'disbursement_no')
  browser.find_element_by_xpath(element).send_keys(no)


def bdx_bord_No(browser,no):
  #border creation, for selecting claim type
  wait_for_overlag(browser)
  app_bord = get_application_url('marc_bord_creation')
  element=get_xpath_by_key(app_bord.application.appId, 'bdx_no')
  browser.find_element_by_xpath(element).clear()
  browser.find_element_by_xpath(element).send_keys(no)



def bdx_Input_Id(CaseIdList,browser,adm_base,post=False):
  app_bord = get_application_url('marc_bord_creation')
  element=get_xpath_by_key(app_bord.application.appId, 'add_record')
  error_xpath=get_xpath_by_key(app_bord.application.appId, 'error_message')
  wait_to_click(browser,element)
  browser.find_element_by_xpath(element).click()
  elementinput=get_xpath_by_key(app_bord.application.appId, 'adm_case')
  if post==True:
   elementinput=get_xpath_by_key(app_bord.application.appId, 'csu_cba_case_id')
  for i in CaseIdList:
    wait_for_overlag(browser)
    browser.find_element_by_xpath(elementinput).clear()
    wait_for_overlag(browser)
    browser.find_element_by_xpath(elementinput).send_keys(i.caseid)
    wait_for_overlag(browser)
    browser.find_element_by_xpath(element).click()
    wait_for_overlag(browser)
    try:
      web_element=browser.find_element_by_xpath(error_xpath)
      error_text=get_text_excluding_children(browser, web_element)
      audit_log("Case id {0} is failed to bordereaux,error message is {1}".format(i.caseid,error_text), "Completed...", adm_base)
    except:
      audit_log("Case id {0} is success to bordereaux".format(i.caseid), "Completed...", adm_base)


def bdx_Date(browser,no):
  #border creation, for selecting Date
  app_bord = get_application_url('marc_bord_creation')
  element=get_xpath_by_key(app_bord.application.appId, 'date_submit')
  browser.find_element_by_xpath(element).send_keys(no)

def get_ClientInsurance_Service(browser):
  #get the client name , not billing name, require wait because come back from other page
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'svc')
  wait_to_load_long(browser,get_xpath_by_key(app_inpaitient.application.appId, 'edit_case'))##specific to wait to load to paitient info,due to cannot seek different elemnet at same time
  wait_to_load_long(browser,element)
  Client=browser.find_element_by_xpath(element).get_attribute("value")
  return Client

def get_OB_Recv_Date(browser):
  #get OB_Recv_Date
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'ws_tab_cba_info_recv_date')
  wait_to_load_long(browser,element)
  OB_Recv_Date=browser.find_element_by_xpath(element).get_attribute("value")
  return OB_Recv_Date

def get_Bill_Reg_Date(browser):
  #get Bill_Reg_Date
  app_inpaitient = get_application_url('marc_search_inpatient')
  element=get_xpath_by_key(app_inpaitient.application.appId, 'ws_tab_cba_info_reg_date')
  wait_to_load_long(browser,element)
  Bill_Reg_Date=browser.find_element_by_xpath(element).get_attribute("value")
  return Bill_Reg_Date

def close_Marc(browser):
  try:
    browser.close()
    browser.quit()
  except Exception as error:
    pass

