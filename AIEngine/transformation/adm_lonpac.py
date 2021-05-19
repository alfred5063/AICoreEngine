#!/usr/bin/python
# FINAL SCRIPT updated as of 12th nov 2019
# Workflow - CBA/ADMISSION
# Version 1
from automagic.marc_adm import *
from transformation.adm_connect_db import *
import automagica
import selenium
from automagica import pyautogui
from selenium.common.exceptions import TimeoutException
from utils.audit_trail import audit_log

import keyboard




def print_First_call(object_list,browser,download_path,adm_base):
  """
  This function automate the process start form the inpaitient search page.

  Arguement is case id , a browser that generated

  End by printing a LONPAC first call
  """
  CBA_Passed_Caseid=[]
  for i in object_list:
    if i.First_call:
      CBA_Passed_Caseid.append(i.caseid)
  if len(CBA_Passed_Caseid)==0:
    return None

  for id in CBA_Passed_Caseid:
    try:
      go_Inpatient_Admission_View(browser)
      key_InID(str(id),browser)
      click_PaitientName(browser)
      click_Edit(browser)
      click_Doc(browser)
      #address=get_Client_Address(browser)
      #NEED TO EDIT TO GET DATA FROM DATABASE
      Cname=get_Client_Name(browser)
      select_First_Call(browser)
      click_None_Doc(browser)

      select_Insurer(browser)
      wait_for_overlag(browser)
      browser.find_element_by_xpath('//*[@id="caseTab:attnTo"]').send_keys(Cname)
      #browser.find_element_by_xpath('//*[@id="caseTab:tabDocFaxEmail"]').send_keys(address)
      browser.find_element_by_xpath('//*[@id="caseTab:tabDocCreateDocButton"]/span[1]').click()
      wait_for_overlag(browser)
      wait_to_click(browser,'//*[@id="caseTab:editDocSaveButton"]/span')
      browser.find_element_by_xpath('//*[@id="caseTab:editDocSaveButton"]/span').click()
      wait_for_overlag(browser)
      wait_to_click(browser,'/html[1]/body[1]/div[1]/div[3]/form[1]/div[3]/div[2]/div[1]/div[1]/div[6]/div[3]/div[2]/div[1]/div[1]/button[1]/span[2]')
      browser.find_element_by_xpath('/html[1]/body[1]/div[1]/div[3]/form[1]/div[3]/div[2]/div[1]/div[1]/div[6]/div[3]/div[2]/div[1]/div[1]/button[1]/span[2]').click()
      Wait(6)
      pyautogui.FAILSAFE = False
      lonpac_path=download_path+'\\'+str(id)+"_FIRSTCALL_doc"

      print(lonpac_path)
      try:
        keyboard.write(lonpac_path)
      except:
        audit_log("keyboard fail, trying another", "Completed...", adm_base)
        try:
          for i in lonpac_path:
            PressKey(key=i)
        except:
          audit_log("keyboard fail, no lonpac file generated", "Completed...", adm_base)
      pyautogui.typewrite(lonpac_path,interval=0.25) 
      try:
        Wait(5)
        PressKey(key="enter")
        Wait(1)
        PressKey(key="left")
        Wait(1)
        PressKey(key="enter")
        Wait(1)
        for object in object_list:
          if id==object.caseid:
            object.First_call_path=str(id)+"_FIRSTCALL_doc"
      except:
        audit_log("file is saved in database", "Completed...", adm_base)
    except TimeoutException:
      id_list.append(id)
      #######3need add unlock
