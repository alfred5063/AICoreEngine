#!/usr/bin/python
# FINAL SCRIPT updated as of 22nd June 2020
# Workflow - CBA/ADMISSION
# Version 1

from automagic.marc_adm import *
from automagica import *
from transformation.adm_connect_db import *
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from transformation.adm_fill_bdx import Check_post_or_type
from utils.audit_trail import audit_log

class adm:
  caseid = 0
  post_main_caseid = 0
  GL_Amount = 0
  GL_Amount_post = 0
  ID_Status = ""
  Status = False
  Post = False
  Date = ""
  Double_Billing = False
  First_call = False
  client_name = ""
  policy_type = ""
  invoice_type = ""
  Bord_Number = ""
  Disburment_No = ""
  att = ""
  add = ""
  bdx_type = ""
  ws_path = ""
  First_call_path = ""
  Ac_ws = ""
  Dr_Fee_Non = ""
  Dr_Fee = ""
  bdx_path = ""
  ack_path = ""
  def __init__(self, caseid):
    self.caseid = caseid


def check_CBA(list_ID, browser, adm_base):
  x = 0
  idtype = ''
  for x in range(len(list_ID)):
    if Check_post_or_type(list_ID[x]) == "POST":
      idtype = 'POST'
      caseid_object_list = []
      for id in list_ID:
        post_main_case_id_list = list(str(post_get_main_case_id(id)) for id in list_ID)
        New_Object = adm(id)
        New_Object.post_main_caseid = (str(post_get_main_case_id(id)))
        caseid_object_list.append(New_Object)
    else:
      idtype = 'CASHLESS'
      caseid_object_list = []
      for id in list_ID:
        New_Object = adm(id)
        New_Object.post_main_caseid = str(id)
        caseid_object_list.append(New_Object)

    if idtype == 'POST':
      for i in caseid_object_list:
        i.Post = True
      object_list = check_CBA_post(caseid_object_list, list_ID, browser, adm_base)
      return object_list
    elif idtype == 'CASHLESS':
      for i in caseid_object_list:
        i.Post = False
      object_list = check_CBA_Cashless(caseid_object_list,list_ID,browser,adm_base)
      return object_list


def check_CBA_Cashless(caseid_object_list, list_ID, browser, adm_base):
  Client_Name = ''
  pol = ''
  counter = 0
  try:
    go_Inpatient_Admission_View(browser)
    audit_log("Admission - Marc login login successfuly", "Completed...", adm_base)
  except:
    audit_log("Admission - Marc failed to login", "Completed...", adm_base)
  for id in list_ID:
    go_Inpatient_Admission_View(browser)
    audit_log("Admission - {} is under CBA checking".format(id), "Completed...", adm_base)
    key_InID(id, browser)

    try:
      click_PaitientName(browser)
      click_Ws(browser)
    except (NoSuchElementException, TimeoutException) as e:
      caseid_object_list[counter].ID_Status = "ID cannot click in marc"
      audit_log("Admission - {} cannot click ID in MARC".format(id), "Completed...", adm_base)
      counter+=1
      continue
    GL_Amount = browser.find_element_by_xpath('/html[1]/body[1]/div[1]/div[3]/form[1]/div[3]/div[2]/div[1]/div[1]/div[4]/table[1]/tbody[1]/tr[8]/td[2]/input[1]').get_attribute("value")
    AC_Amount = browser.find_element_by_xpath('/html[1]/body[1]/div[1]/div[3]/form[1]/div[3]/div[2]/div[1]/div[1]/div[4]/table[1]/tbody[1]/tr[1]/td[4]/input[1]').get_attribute("value")
    OB_Amount = browser.find_element_by_xpath('/html[1]/body[1]/div[1]/div[3]/form[1]/div[3]/div[2]/div[1]/div[1]/div[4]/div[1]/div[1]/div[2]/table[1]/tbody[1]/tr[1]/td[6]/input[1]').get_attribute("value")
    discount_Amount = browser.find_element_by_xpath('/html[1]/body[1]/div[1]/div[3]/form[1]/div[3]/div[2]/div[1]/div[1]/div[4]/table[1]/tbody[1]/tr[2]/td[4]/input[1]').get_attribute("value")

    Client_Name = get_Client_Name(browser)
    pol = get_Pol_Type(browser)
    caseid_object_list[counter].client_name = Client_Name
    caseid_object_list[counter].policy_type=pol
    if check_Amount(GL_Amount,OB_Amount):
      caseid_object_list[counter].GL_Amount=GL_Amount
      if check_Status(browser,post=Check_post_or_type(id)):
        if check_Doc(browser,id):
          audit_log("Admission - {} CBA checking is passed,and ready to bdx".format(id), "Completed...", adm_base)
          caseid_object_list[counter].ID_Status="Pass"
          caseid_object_list[counter].Status=True
          Client_Name= get_Client_Name(browser)
          if Client_Name.upper().find('LONPAC') != (-1) or  Client_Name.upper().find('HONG LEONG') != (-1):
            caseid_object_list[counter].First_call=True
          try:
            click_Edit(browser)
            print("bdx")
            change_to_BdxReady(browser)
            print("unlock")
            click_unlock(browser)
            counter+=1
            
          except:
            caseid_object_list[counter].ID_Status="Marc Case id Cannot be Edited"
            audit_log("{} cannot be edited in MARC/is locked".format(id), "Completed...", adm_base)
            caseid_object_list[counter].Status=False
            counter+=1
        else:
          audit_log("{} have no document scanned in Marc".format(id), "Completed...", adm_base)
          caseid_object_list[counter].ID_Status="Case id have no document scanned in Marc"
          counter+=1
      else:
        audit_log("{} CBA status is not yet ready for bdx".format(id), "Completed...", adm_base)
        caseid_object_list[counter].ID_Status="CBA Status not Completed"
        counter+=1
    else:
      audit_log("{} Amount different with OB And GL".format(id), "Completed...", adm_base)
      caseid_object_list[counter].ID_Status="{} Amount different with OB And GL. ".format(id)+"OB={},GL={}".format(OB_Amount,GL_Amount)
      counter+=1
  return caseid_object_list

def check_CBA_post(caseid_object_list, list_ID_post, browser, adm_base):
  counter = 0
  pol = ""
  go_Inpatient_Admission_Post_View(browser)
  for id in list_ID_post:
    go_Inpatient_Admission_Post_View(browser)
    audit_log("{} is under CBA post checking".format(id), "Completed...", adm_base)
    key_InID(id, browser)

    try:
      click_PaitientName_post(browser)
      click_Ws(browser)
    except (NoSuchElementException,TimeoutException) as e:
      caseid_object_list[counter].ID_Status="ID cannot click in marc in POST"
      audit_log("{} cannot click ID in MARC".format(id), "Completed...", adm_base)
      caseid_object_list[counter].Status=False
      counter+=1
      continue
    GL_Amount=browser.find_element_by_xpath('/html[1]/body[1]/div[1]/div[3]/form[1]/div[4]/div[2]/div[1]/div[1]/div[4]/table[1]/tbody[1]/tr[10]/td[2]/input[1]').get_attribute("value")
    AC_Amount=browser.find_element_by_xpath('/html[1]/body[1]/div[1]/div[3]/form[1]/div[4]/div[2]/div[1]/div[1]/div[4]/table[1]/tbody[1]/tr[1]/td[4]/input[1]').get_attribute("value")
    discount_Amount=browser.find_element_by_xpath('/html[1]/body[1]/div[1]/div[3]/form[1]/div[4]/div[2]/div[1]/div[1]/div[4]/table[1]/tbody[1]/tr[2]/td[4]/input[1]').get_attribute("value")
    Client_Name=get_Client_Name_post(browser)
    pol=get_Pol_Type_post(browser)
    caseid_object_list[counter].client_name=Client_Name
    caseid_object_list[counter].policy_type=pol
    if True:
      caseid_object_list[counter].GL_Amount_post=GL_Amount
      if check_Status_post(browser):
        caseid_object_list[counter].Status=True
        caseid_object_list[counter].ID_Status="Pass"
        audit_log("{} CBA checking is passed,and ready to bdx".format(id), "Completed...", adm_base)
        counter+=1
      else:
        audit_log("{} CBA status is not yet ready for bdx".format(id), "Completed...", adm_base)
        caseid_object_list[counter].ID_Status="CBA Status in POST not Completed"
        caseid_object_list[counter].Status=False
        counter+=1
    else:
      audit_log("{} Amount different with discount amount And AC amount in POST.format(id)", "Completed...", adm_base)
      caseid_object_list[counter].ID_Status="Amount different with discount amount And AC amount in POST"
      caseid_object_list[counter].Status=False
      counter+=1
  return caseid_object_list


def change_CBA_Status(browser, object_list):
  new_list = []
  for i in object_list:
    if i.ID_Status == "Pass":
      new_list.append(i)
      Wait(5)
      if i.bdx_type == "Post":
        go_Inpatient_Admission_Post_View(browser)
        time.sleep(5)
        key_InID(i.caseid, browser)
        time.sleep(5)
        click_PaitientName(browser)
        time.sleep(5)
        element_edit = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/a[2]'
        edit_record = browser.find_element_by_xpath(element_edit)
        ActionChains(browser).move_to_element(edit_record).click(edit_record).perform()
        time.sleep(5)
        element_bdx = '/html/body/div[1]/div[3]/form/div[1]/div[2]/div[1]/table/tbody/tr/td[6]/select'
        Select(browser.find_element_by_xpath(element_bdx)).select_by_visible_text("CBA - Ready for Bdx")
        time.sleep(5)
        element_save = '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/a[2]'
        save_record = browser.find_element_by_xpath(element_save)
        ActionChains(browser).move_to_element(save_record).click(save_record).perform()
        time.sleep(5)
      else:
        go_Inpatient_Admission_View(browser)
        time.sleep(5)
        key_InID(i.caseid, browser)
        time.sleep(5)
        click_PaitientName(browser)
        time.sleep(5)
        click_Edit(browser)
        time.sleep(5)
        change_to_BdxReady(browser)
        time.sleep(5)
        click_Save(browser)
        time.sleep(5)
    else:
      continue
  return new_list
