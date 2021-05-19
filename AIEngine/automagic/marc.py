# Declare Python libraries needed for this script
#!/usr/bin/python
# FINAL SCRIPT updated as of 12th March 2021

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

current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)

config = read_db_config(dti_path+r'\config.ini', 'debug')
mode = config['mode']

def initBrowser(base, url=None, profile=None, username=None, password=None):
  
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
    
    audit_log('Login marc','Trying to load the login page at {0}'.format(app.application.appUrl), base)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'username')).send_keys(user)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'password')).send_keys(pwd)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'login_button')).click()
    wait_to_load(browser, "/html[1]/body[1]/div[1]/div[3]/div[1]/div[1]/div[2]/form[1]/button[1]", 60)
    valid_login = False
    try:
      login_status = browser.find_element_by_xpath("/html[1]/body[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[1]").text()
      if login_status == "Login/Password incorrect.":
        logging('Invalid login/password', 'login id: {0}'.format(user), base, status=mode)
        valid_login = False
    except Exception as error:
      if "Message: Unable to locate element: /html[1]/body[1]/div[1]/div[1]/div[2]/form[1]/div[1]/div[1]" in str(error): 
        print('login marc successful.')
        valid_login = True
        audit_log('Import to MARC', 'Login to MARC successfully', base, status=mode)
      else:
        print('error: {0}'.format(error))

    return browser, valid_login
   
  except Exception as error:
    logging('login issue due to invalid password', error, base)
    browser.close()

def upload_import(base, browser, filepath, debugmode=False):
  #pyautogui.FAILSAFE= True
  try:
    mode = debugmode
    audit_log('Navigate to import site', 'Completed...', base, status=mode)
    app = get_application_url('marc_data_import')
    browser.get(app.application.appUrl)
    audit_log('Import to MARC', 'Upload file started', base, status=mode)
    audit_log('Select file', 'Completed...', base, status=mode)
    
    uploadElement = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'file_select'))
    uploadElement.send_keys(filepath)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId,'file_upload')).click()
    time.sleep(5)
    #browser.switch_to_active_element().send_keys(filepath)
    #PressHotkey('alt', 'n')
    #Type(filepath, interval_seconds=0.01)
    #PressHotkey('alt', 'o')

    import_status = ''
    audit_log('File upload progress', 'Completed...', base, status=mode)
    browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'file_upload')).click()
    wait_to_load(browser, get_xpath_by_key(app.application.appId, 'upload_import_status'),60)
    batch_num = browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'get_batch_num')).get_attribute("innerHTML")
    import_status = batch_num.split(' | ')
    audit_log('Import to MARC', 'Upload file successfully', base, status=mode)
    audit_log('File upload completed', 'Completed...', base, status=mode)
  except Exception as error:
    print(error)
    logging('upload_import', error, base, status=mode)
  return browser, import_status

def wait_to_load(browser,element, second):
  #wait for element to be show,suggested use in loading page for browser
  #try:
  wait=WebDriverWait(browser,second)
  wait.until(EC.visibility_of_element_located((By.XPATH, element)))

def wait_for_overlag(browser):
  #Wait(5)
  #wait for the webpage to load the overlay, overlay is the delay that make mouse cannot click
  wait_overlag=WebDriverWait(browser,20)
  #wait_overlag.until(EC.invisibility_of_element_located((By.XPATH, "/html[1]/body[1]/div[3]")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2626")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt103")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2588")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2315")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2810")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2599")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.ID,"j_idt2868")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.CLASS_NAME,'ui-widget-overlay ui-dialog-mask')))
  #wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[6]")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[22]")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[1]/div[6]")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[16]")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[4]/div[1]/div[1]")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"/html[1]/body[1]/div[4]/div[1]/div[1]")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"//div[@class='ui-growl-item']")))
  #wait_overlag.until(EC.invisibility_of_element_located((By.XPATH,"//div[@class='three-quarters']")))
  wait_overlag.until(EC.invisibility_of_element_located((By.CLASS_NAME,'three-quarters')))

def wait_to_click(browser,element):
  #wait for element to be show,suggested use in JAVAscript loading page for browser
  wait=WebDriverWait(browser,3)
  wait.until(EC.element_to_be_clickable((By.XPATH, element)))

def get_xpath_from_list(xpathlist, key):
  values =  xpathlist[xpathlist["key"]==key].xpath
  value = ""
  for item in values:
    value = item
 
  return value

def search_member(base, browser, search_criteria, row, marc_search_page, app1, marc_update_member):
  audit_log('search_member', 'Completed...MARC search member started', base, status=mode)
  print('MARC search member started')
  
  browser.get(app1)

  #split fields into proper dictionary
  wait_to_load(browser, get_xpath_from_list(marc_search_page, 'policy num'), 60)
  fields = [i.strip(" ") for i in search_criteria.split(";")]
  for field in fields:
    browser.find_element_by_xpath(get_xpath_from_list(marc_search_page, field)).send_keys(row[field])
    continue
  #click search to find record for membership
  browser.find_element_by_xpath(get_xpath_from_list(marc_search_page, 'search')).click()

  #wait finish loading
  wait_for_overlag(browser)
  #Wait(10)
  #browser.find_element_by_xpath(get_xpath_by_key(app.application.appId, 'loading_indicator'))
  elems = browser.find_elements_by_xpath(get_xpath_from_list(marc_search_page, 'edit_record'))
  
  id_mm = None
  if len(elems) >= 1:
    i = 1
    for elem in elems:
      try:
        name = browser.find_element_by_xpath(get_xpath_from_list(marc_search_page, 'red_search_member').format(i))
      except:
        browser.find_element_by_xpath(get_xpath_from_list(marc_search_page, 'search_valid_member').format(i)).click()
        wait_to_load(browser, get_xpath_from_list(marc_update_member, 'id_mm'), 60)
        id_mm = browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'id_mm')).get_attribute("value")
      i=i+1
  else:
    elems[0].click()
    wait_to_load(browser, get_xpath_from_list(marc_update_member, 'id_mm'), 60)
    id_mm = browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'id_mm')).get_attribute("value")
    
  audit_log('search_member', 'Completed...MARC search member ended', base, status=mode)

  print('MARC search member ended')
  return browser, row, id_mm

def update_member(base, browser, row, marc_update_member, appUrl):
  
  audit_log('update_member', 'Completed...MARC update member started', base, status=mode)
  print('MARC update member started')
  fields_to_update = (row["fields_update"].split(";")[row["action"].split(";").index("update_fields")]).split("_")[0:]
  try:
    print(row["status"])
  except Exception as error:
    print('Define status column.')
    row["status"] = ""

  if len(fields_to_update) == 0:
    print('No field to update.')
    row["status"] = row["status"]+"<No field update.>"
  else:
    #app = get_application_url('marc_setup_member_policy')
    flag = True
    browser.get(appUrl)
    id_mm = browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'id_mm')).get_attribute("value")
    for field in fields_to_update:
      if field == "joindate":
        get_first_join_date = get_xpath_from_list(marc_update_member, 'first join date')
        browser.find_element_by_xpath(get_first_join_date).clear()
        browser.find_element_by_xpath(get_first_join_date).send_keys(row["first join date"])
        continue
      elif field == "planattachdate":
        get_plan_attach_date = get_xpath_from_list(marc_update_member, "plan attach date")
        browser.find_element_by_xpath(get_plan_attach_date).clear()
        browser.find_element_by_xpath(get_plan_attach_date).send_keys(row["plan attach date"])
        continue
      elif field == "relationship":
        get_relationship = get_xpath_from_list(marc_update_member, "relationship")
        #browser.find_element_by_xpath(get_relationship).clear()
        elem = Select(
                browser.find_element_by_xpath(get_relationship)
            )
        options = {
                item.get_attribute("value").lower(): item.text
                for item in elem.options
                if 'select one' not in item.text
            }
        if row["relationship"].lower() in options.keys():
                elem.select_by_visible_text(
                    options[row["relationship"].lower()]
                )
        else:
              error = "Relationship invalid value, can't select relationship."
              print(error)
              logging("update relationship", error, base)
              row["error"] = row["error"]+"<Relationship error: {0}>".format(error)
        continue
      elif field == "employee":
        get_employee_id = get_xpath_from_list(marc_update_member, "employee id")
        browser.find_element_by_xpath(get_employee_id).clear()
        browser.find_element_by_xpath(get_employee_id).send_keys(row["employee id"])
        continue
      elif field == "extref":
        get_ext_ref = get_xpath_from_list(marc_update_member, "extref")
        browser.find_element_by_xpath(get_ext_ref).clear()
        browser.find_element_by_xpath(get_ext_ref).send_keys(row["external ref id  aka client "])
        continue
      elif field == "plan":
        browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, "plan_button")).click()
        wait_for_overlag(browser)
        if row["external plan code"] == "":
          browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, "plan_int_ref")).send_keys(row["internal plan code id"])
        else:
          browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, "plan_ext_ref")).send_keys(row["external plan code"])

        browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'plan_search')).click()
        wait_for_overlag(browser)
        #wait_to_load(browser, '/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/form[1]/div[2]/div[2]/table[1]/tbody[1]/tr[1]',60)
        check_record = browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'select_plan'))
        if check_record.text =="No records found.":
          if row["internal plan code id"] != "":
            browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, "plan_ext_ref")).clear()
            browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, "plan_int_ref")).send_keys(row["internal plan code id"])
            browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'plan_search')).click()
            check_internal_record = browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'select_plan'))
            if check_internal_record == "No records found.":
              browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'plan_close')).click()
              wait_for_overlag(browser)
              row["status"] = row["status"]+"<Plan - No plan exist>"
              row["error"] = row["error"]+"<Internal plan value not found in MARC: {0}>".format(row["internal plan code id"])
              flag=False
            else:
              select_plan = browser.find_element_by_xpath('{0}/a[1]'.format(get_xpath_from_list(marc_update_member, 'select_plan')))
              select_plan.click()
              wait_for_overlag(browser)
          else:
            browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'plan_close')).click()
            wait_for_overlag(browser)
            row["status"] = row["status"]+"<Plan - No plan exist>"
            row["error"] = row["error"]+"<External plan value not found in MARC: {0}>".format(row["external plan code"])
            flag=False
        else:
          total_plan = browser.find_elements_by_xpath(get_xpath_from_list(marc_update_member, 'select_plan'))
       
          if len(total_plan) > 1:
            browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'plan_close')).click()
            wait_for_overlag(browser)
            row["status"] = row["status"]+"<Plan - more than one plan>"
            row["error"] = row["error"]+"<Plan - plan search is more than one plan, rpa can't not recognize.>"
            flag=False
          elif len(total_plan) == 1:
            select_plan = browser.find_element_by_xpath('{0}/a[1]'.format(get_xpath_from_list(marc_update_member, 'select_plan')))
            select_plan.click()
            wait_for_overlag(browser)
          else:
            row["status"] = row["status"]+"<Plan - unknown issue>"
            row["error"] = row["error"]+"<Plan - plan search is issue, rpa can't not recognize.>"
            browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'plan_close')).click()
            wait_for_overlag(browser)
            flag=False
          continue
      elif field == "principal":
        if row["relationship"]== "P" or row["relationship"] == "Principal":
          row["status"] = row["status"]+"<Relationship is principal. Doesn't need to update this field.>"
        else:
          browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, "principal_click")).click()
          #wait_to_load(browser, "/html[1]/body[1]/div[1]/div[3]/div[4]/div[1]/span[1]", 60)
          wait_for_overlag(browser)
          if row["principal nric"]=="":
            if row["principal other ic"]=="":
              if row["principal ext ref id (aka client)"]=="":
                row["status"] = row["status"]+"<Can't find principal details.>"
              else:
                browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, "principal_extref")).send_keys(row["principal ext ref id (aka client)"])
            else:
              browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, "principal_otheric")).send_keys(row["principal other ic"])
          else:
            #print(row['principal nric'])
            browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, "principal_nric")).send_keys(row["principal nric"])

          #browser.find_element_by_xpath(get_xpath_from_list(marc_update_member,"principal_expiry_date")).send_keys(row["plan expiry date"])
          browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'principal_search')).click()
          wait_for_overlag(browser)
          check_principal_selection = browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'principal_select'))
          if check_principal_selection.text == "No records found.":
            row["status"] = row["status"]+"<Principal not exist>"
            row["error"] = row["error"]+"<Principal is not exist for this field, rpa can't update.>"
            browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'principal_close')).click()
            wait_for_overlag(browser)
            
          else:
            total_principal_selection = browser.find_elements_by_xpath(get_xpath_from_list(marc_update_member, 'principal_select'))
            if len(total_principal_selection) > 1:
              row["status"] = row["status"]+"<Principal more than one exist>"
              row["error"] = row["error"]+"<Principal is more than one, rpa can't not recognize.>"
              browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'principal_close')).click()
              wait_for_overlag(browser)
              flag=False
            elif len(total_principal_selection) == 0:
              row["status"] = row["status"]+"<Principal not exist>"
              row["error"] = row["error"]+"<Principal is not exist for this field, rpa can't update.>"
              browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'principal_close')).click()
              wait_for_overlag(browser)
              flag=False
            elif len(total_principal_selection)==1:
              select_principal = browser.find_element_by_xpath('{0}/a[1]'.format(get_xpath_from_list(marc_update_member, 'principal_select')))
              select_principal.click()
              wait_for_overlag(browser)
            else:
              row["status"] = row["status"]+"<Principal - unknown issue>"
              row["error"] = row["error"]+"<Principal - plan search is issue, rpa can't not recognize.>"
              browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'principal_close')).click()
              wait_for_overlag(browser)
              flag=False
        continue

    browser.find_element_by_xpath(get_xpath_from_list(marc_update_member, 'update button')).click()
    #wait_to_load(browser, "/html[1]/body[1]/div[3]/div[1]/div[1]/div[2]/p[1]", 60)
    wait_for_overlag(browser)

    if flag == True:
      row["record_status"]= "Successful"
    else:
      row["record_status"] = "Unsuccessful"
    
  audit_log('update_member', 'Completed...MARC update member ended', base, status=mode)
  print('MARC update member ended')
  return browser, row, flag, id_mm

def update_member_details(base, browser, row, marc_update_member_details, url, imm_id, marc_update_member_details_edit):
  audit_log('update_member_details', 'Completed...MARC update member details started', base, status=mode)
  print('MARC update member details started')
  flag = True
  fields_to_update = (row["fields_update"].split(";")[row["action"].split(";").index("update_member")]).split("_")[0:]
  if imm_id != "":
    if len(fields_to_update) == 0:
      print('No field to update.')
      row["status"] = row["status"]+"<No member details field update.>"
    else:
      #app = get_application_url('marc_setup_member_policy')
      flag = True
      browser.get(url)
      wait_for_overlag(browser)
      browser.find_element_by_xpath(get_xpath_from_list(marc_update_member_details, "search_member_imm_id")).send_keys(imm_id)
      browser.find_element_by_xpath(get_xpath_from_list(marc_update_member_details, "search_button")).click()
      print('search result completed.')
      #wait finish loading
      wait_for_overlag(browser)
      
      elems = browser.find_elements_by_xpath(get_xpath_from_list(marc_update_member_details, 'edit_record'))
      if len(elems) != 1:
        row["status"] = row["status"]+"<Found more than one membership>"
        row["error"] = row["error"]+"<Found more than one membership, rpa can't not recognize.>"
        flag=False
      else:
        elems[0].click()
        Wait(2)
        try:
          for field in fields_to_update:
            if field == "name":
              update_name = get_xpath_from_list(marc_update_member_details_edit, "name")
              browser.find_element_by_xpath(update_name).clear()
              browser.find_element_by_xpath(update_name).send_keys(row["member full name"])
            elif field == "nric":
              if row["nric"] != "":
                update_nric = get_xpath_from_list(marc_update_member_details_edit, "nric")
                browser.find_element_by_xpath(update_nric).clear()
                print(row['nric'])
                browser.find_element_by_xpath(update_nric).send_keys(row["nric"])
            elif field == "otheric":
              if row["otheric"] != "":
                update_otheric = get_xpath_from_list(marc_update_member_details_edit, "otheric")
                browser.find_element_by_xpath(update_otheric).clear()
                browser.find_element_by_xpath(update_otheric).send_keys(row["otheric"])
            elif field == "dob":
                update_dob = get_xpath_from_list(marc_update_member_details_edit, "dob")
                browser.find_element_by_xpath(update_dob).clear()
                browser.find_element_by_xpath(update_dob).send_keys(row["dob"])
            flag = True
        except Exception as error:
          flag = False
          logging("update_member_details", error, base)

        wait_for_overlag(browser)
        browser.find_element_by_xpath(get_xpath_from_list(marc_update_member_details_edit, "update")).click()
        #wait_to_load(browser, "/html[1]/body[1]/div[1]/div[3]/form[1]/div[1]/div[1]/div[1]/ul[1]/li[1]/span[1]", 60)
        wait_for_overlag(browser)

      print('MARC update member details ended')
      audit_log('update_member_details', 'Completed...MARC update member details ended', base, status=mode)
  else:
    print('Imm Id not found.')
    row["status"] = row["status"]+"<Imm Id not found, search member id required.>"
    row["error"] = row["error"]+"<Imm Id not found, search member id required.>"
    #for field in fields_to_update:
  return browser, row, flag
