#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - CBA/ADMISSION

# Declare Python libraries needed for this script
import re
import time
import os
import automagica
import json
import pandas as pd
import ast
import datetime as dt
from automagic.marc_adm import *
from transformation.adm_connect_db import *
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from datetime import date, datetime, timedelta
from transformation.adm_bdx_manipulation import Get_Address
from connector.connector import MySqlConnector
from transformation.adm_bdx_manipulation import get_Lastest_File
from utils.audit_trail import audit_log

class invoice_class():
  b2b = False
  cashless = False
  post = False
  excess = False

def Bord_Listing(id, invoice_type, adm_base):

  if "post" in invoice_type.lower():
    file_name_convention = "(P)"
  else:
    if "excess" in invoice_type.lower():
      file_name_convention = "(E)"
    else:
      if "b2b" in invoice_type.lower():
        file_name_convention = "(B)"
      else:
        file_name_convention = "(C)"
  
  Bord_No = None
  i = 1
  try:
    Bord_file = get_lastest_BordNum(id, invoice_type)
    while Bord_No == None:
      Bord_No = (json.loads(Bord_file[0][0])['Data'][-i]['Disbursement Listing'])
      i+=1
  except:
    Bord_No = "00/00"
  Bord_No = Bord_No.split("/")
  Number = str(int(Bord_No[0])+1)+"/"+str(datetime.now().year)[-2:]+file_name_convention
  return Number


def Check_post_or_type(caseid):
  conn = MySqlConnector()
  cur = conn.cursor()

  # Calling Stored Procedure
  parameters = [str(caseid),]
  stored_proc = cur.callproc("query_marc_for_inpatient_validity", parameters)
  for i in cur.stored_results():
    results = i.fetchall()
  fetched_result_adm = pd.DataFrame(results)

  stored_proc = cur.callproc("query_marc_for_outpatient_validity", parameters)
  for i in cur.stored_results():
    results = i.fetchall()
  fetched_result_post = pd.DataFrame(results)

  cur.close()
  conn.close()

  # First Check
  if fetched_result_adm.empty != True and fetched_result_post.empty != True:
    if str(fetched_result_post[0][0]) != str(caseid):
      type = "POST"
    elif str(fetched_result_adm[0][0]) == str(caseid):
      type = "CASHLESS"
  elif fetched_result_adm.empty != True and fetched_result_post.empty == True:
    type = "CASHLESS"
  elif fetched_result_adm.empty == True and fetched_result_post.empty != True:
    type = "POST"
  elif fetched_result_adm.empty == True and fetched_result_post.empty == True:
    type = "ERROR"

  return type


def get_client_db_variables(browser, object_list, adm_base):\
  #casetype = Check_post_or_type(object_list[0].caseid)
  first_adm_obj = object_list[0]
  first_case_id = first_adm_obj.caseid
  invoice_char = invoice_class()
  if first_adm_obj.Post == True:
    type = 'POST'
    invoice_char.post = True
    go_Inpatient_Admission_Post_View(browser)
    key_InID(first_case_id,browser)
    click_PaitientName_post(browser)
    comp_pay = get_comp_pay_post(browser)
    if check_double_billing_post(browser) == True:
      invoice_char.excess = True
    elif check_double_billing_post(browser) == False:
      if Check_B2B(browser, comp_pay, type) != True:
        invoice_char.cashless = True
      else:
        invoice_char.b2b = True
  elif first_adm_obj.Post != True:
    type = 'ADM'
    go_Inpatient_Admission_View(browser)
    key_InID(first_case_id,browser)
    click_PaitientName(browser)
    click_Edit(browser)
    comp_pay = get_comp_pay(browser)
    if check_double_billing(browser) == True:
      invoice_char.excess = True
    elif check_double_billing(browser) == False:
      if Check_B2B(browser, comp_pay, type) != True:
        invoice_char.cashless = True
      else:
        invoice_char.b2b = True
    client_name = get_Client_Name(browser)
    click_unlock(browser)

  print(invoice_char.post)
  print(invoice_char.excess)
  print(invoice_char.b2b)
  print(invoice_char.cashless)
  return invoice_char


def get_specific_invoice_type(all_invoice_type, invoice_char, adm_base):
  try:
    checker = []
    temp_list = []
    enlisted = []
    if invoice_char.post:
      checker.append("post")
    if invoice_char.cashless:
      checker.append("cashless")
    if invoice_char.b2b:
      checker.append("b2b")
    if invoice_char.excess:
      checker.append("excess")
    if len(checker) == 1:
      for i in all_invoice_type:
        if checker[0].lower() in i[0].lower() and checker[0].lower() == i[0].lower():
          enlisted.append(i[0])
          enlisted = list(set(enlisted))
      return enlisted
    else:
      for i in all_invoice_type:
        if checker[0].lower() in i[0].lower():
          temp_list.append(i[0])
      for i in temp_list:
        if checker[1].lower() == "cashless":
          if "b2b" not in i.lower():
            enlisted.append(i)
        else:
          if checker[1].lower() in i.lower():
            enlisted.append(i)
      if len(enlisted) != 0:
        return enlisted
      else:
        for i in all_invoice_type:
          if "post" == i[0].lower().strip():
            return i[0]
  except:
    audit_log("Admission - WRONG Invoice type is taken {}".format(all_invoice_type[0]), "Completed...", adm_base)
    return all_invoice_type


def get_all_db_invoice_type(client_id, adm_base):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry = '''select [invoice_type] from cba.client_master where [client_id] = %s;'''
  cursor.execute(qry, (str(client_id)))
  records = cursor.fetchall()

  # Quick fix
  if records == []:
    records = [('B2B',), ('B2B',), ('Cashless',), ('Cashless',), ('Post B2B (full)',), ('NO BORD',), ('Post B2B (Full)',), ('Post (Insurance Only)',), ('Cashless to Allianz ',), ('Post to Allianz ',), ("Excess to Gas M'sia",), ("Post Excess to Gas M'sia",), ('Cashless',), ('Post B2B',), ('Cashless',), ('POST FULL B2B',), ('POST ASTRO',), ('Reimbursement ',), ('Cashless',), ('B2B',), ('Post',), ('Cashless',), ('Cashless',), ('Cashless Excess-Leader Cable',), ('Cashless',), ('POST ',), ('NO BORD',), ('Cashless',), ('Cashless',), ('B2B',), ('Post B2B',), ('Cashless',), ('B2B',), ('Post',), ('Cashless',), ('B2B',), ('Post',), ('Cashless',), ('B2B',), ('Post',), ('Post B2B',), ('Cashless Excess-MDEC',), ('Post Excess-MDEC ',), ('Cashless',), ('Cashless',), ('Post',), ('Excess',), ('Post Excess',), ('DCA',), ('Cashless',), ('B2B',), ('POST - B2B',), ('Reimbursement ',), ('Cashless',), ('B2B',), ('Post',), ('Cashless',), ('B2B',), ('POST TOK',), ('POST B2B',), ('AEON CO EXCESS',), ('AEON CO POST EXCESS',), ('AEON INDEX LIVING EXCESS',), ('AEON INDEX LIVING POST EXCESS',), ('AEON GLOBAL POST EXCESS',), ('CASHLESS',), ('Cashless',), ('B2B',), ('Post B2B',), ('Cashless',), ('B2B',), ('MMP',), ('Gladiator',), ('Post B2B',), ('MSIG ABOVE 10K ',), ('MSIG BELOW 10K',), ('B2B',), ('MSIG Post B2B',), ('MSIG Post Insurable',), ('Cashless',), ('B2B',), ('Post',), ('Cashless',), ('B2B',), ('Post',), ('Cashless',), ('B2B',), ('Post',), ('Cashless',), ('B2B',), ('Post B2B',), ('Post Excess-Leader Cable',), ('Cashless excess-Leader Universa',), ('Post excess-Leader universal',), ('Cashless Excess-Universal Cable',), ('Post Excess-Universal Cable',), ('Cashless B2B',), ('Post B2B',), ('Excess',), ('Cashless Excess-Lite Kabel',), ('Post Excess-Lite Kabel',), ('Cashless Excess-HNG Capital',), ('Post Excess-HNG Capital',), ('Cashless',), ('Post',), ('PostHR B2B',), ('Reimbursement ',), ('Cashless',), ('B2B',), ('Post',), ('Cashless',), ('B2B',), ('Post',), ('Cashless',), ('DCA',), ('Cashless',), ('Cashless',), ('Cashless',), ('B2B',), ('POST',), ('Cashless',), ('B2B',), ('Post B2B',), ('Cashless',), ('POST ',), ('NO BORD',), ('Cashless',), ('B2B',), ('Post',), ('Post Excess',), ('Cashless',), ('Cashless',), ('B2B',), ('Post',), ('Reimbursement',), ('Reimbursement',), ('DCA',), ('Cashless',), ('B2B',), ('Post B2B',)]
  else:
    records = records

  cursor.close()
  connector.close()
  return records

def hns_newyear_checking(hns):
  now = datetime.now()
  this_year = (str(now.year))[-1]
  dis_year = (str(hns.split('-')[0]))[-1]
  if this_year == dis_year:
    return hns
  else:
    Disburment_No = "H&S"+str(now.year)[2:4] + "-" + "0001"
    print("Disburment_No is %s" % Disburment_No)
    return Disburment_No



def For_Preview(browser, object_list, adm_base):
  """
  put passed case_id into the position number 1
  """
  invalid = True
  for adm_obj in object_list:
    if adm_obj.ID_Status == "Pass":
      invalid = False
      c = adm_obj
      object_list.remove(c)
      object_list.insert(0, c)

  if invalid:
    return 'Cashless', object_list, []

  Insurance = object_list[0].client_name
  Double_Billing = False
  for i in object_list:
    if i.ID_Status == "Pass":
      pass
    elif i == object_list[-1] and i.ID_Status != "Pass":
      enlisted = []
      return enlisted, object_list, Double_Billing #no case id, no need double billing

  Disburment_No = None
  now = datetime.now()
  dis_no_counter = 1
  Dis_file = get_lastest_DirNo()[0][0]
  Disburment_No = json.loads(Dis_file)["Data"][-1]["Disbursement Claims No"]
  Disburment_Number = int(Disburment_No.split('-')[1])
  Disburment_Number+=1
  Disburment_No = "H&S" + str(now.year)[2:4] + "-" + str(Disburment_Number).zfill(5)

  named_tuple = time.localtime()
  Date = time.strftime("%d-%m-%Y", named_tuple)
  att, add, Cliend_id, enlisted = Get_Address(object_list, adm_base)
  invoice_obj = get_client_db_variables(browser, object_list, adm_base)
  all_invoice_type = get_all_db_invoice_type(Cliend_id, adm_base)
  bdx_type = get_specific_invoice_type(all_invoice_type, invoice_obj, adm_base)[0]
  audit_log("Invoice type is {}".format(bdx_type), "Completed...", adm_base)
  try:
    Bord_Number = Bord_Listing(Cliend_id, bdx_type, adm_base)
  except:
    audit_log("Cannot found invoice_type in database, change to Cashless automatically ", "Completed...", adm_base)
    bdx_type = "Cashless"
    Bord_Number = Bord_Listing(Cliend_id, bdx_type, adm_base)

  for i in (object_list):
    if i.ID_Status == "Pass":
      i.Bord_Number = (Bord_Number)
      i.Disburment_No = (Disburment_No)
      i.Date = (Date)
      i.att = (att)
      i.add = (add)
      i.bdx_type = (bdx_type)
      i.Double_Billing = Double_Billing

  return bdx_type, object_list, enlisted


def preview_df(object_list):
  if object_list[0].Post:
    empty = True
    df = pd.DataFrame(columns=list(["Case_ID","ID_Status","Bord_Number","Disbursement_No","Address","Attention","Case_Type","Date","GL_Amount","GL_Amount_Post","Main_Case_ID"]))
    counter = 0
    for i in object_list:
      if i.ID_Status == "Pass":
        empty = False
      df.loc[counter] = [i.caseid, i.ID_Status, i.Bord_Number.replace("/","-"), i.Disburment_No, i.add, i.att, i.bdx_type, i.Date.replace("/","-"), str(i.GL_Amount), str(i.GL_Amount_post), i.post_main_caseid]
      counter+=1
  else:
    empty = True
    df = pd.DataFrame(columns = list(["Case_ID", "ID_Status", "Bord_Number", "Disbursement_No", "Address", "Attention", "Case_Type", "Date", "GL_Amount"]))
    counter = 0
    for i in object_list:
      if i.ID_Status == "Pass":
        empty = False
      df.loc[counter] = [i.caseid, i.ID_Status, i.Bord_Number.replace("/","-"), i.Disburment_No, i.add, i.att, i.bdx_type, i.Date.replace("/","-"), str(i.GL_Amount)]
      counter+=1

  x = df.to_json(orient = 'index')
  json_ob = json.loads(x)
  return json_ob, empty

def get_Date():
  """
  This function get today date and compare with calender event in database
  return a formatted date to bdx 
  """
  Date = dt.date.today() + timedelta(days = 1)
  dateList = []
  for i in get_EventDate():
    temp = list(i)
    dateList.append(temp[0])
  while Date in dateList:
    Date = Date+timedelta(days = 1)
  formatted_Date = Date.strftime("%d/%m/%Y")
  return formatted_Date


def fill_Bdx(browser, object_list, adm_base):
  """
  main function object_list
  this function accept one or more case id in list and automate it to bdx fill in and click save
  """
  i = 0
  while i < len(object_list):
    if object_list[i].Bord_Number is not "":
      i+=1
    else:
      break

  post = False
  j = 0
  for j in range(len(object_list)):
    if "post" in object_list[j].bdx_type.lower():
      post = True

    if post:
      go_Inpatient_Admission_Post_View(browser)
      key_InID(object_list[j].caseid, browser)
      click_PaitientName_post(browser)
      Policy = get_Pol_Type_post(browser)
      Client = object_list[j].client_name
      Client_Short = Client.split(" ")[0]
      go_Bordereaux_Add_New(browser)
      find_client(browser, Client)
      bdx_Post_Cashless_Admission(browser)

      # Get the Discharge Date from MARC?
    else:
      go_Inpatient_Admission_View(browser)
      key_InID(object_list[0].caseid, browser)
      click_PaitientName(browser)
      Policy = get_Pol_Type(browser)
      Client = object_list[0].client_name
      Client_Short = Client.split(" ")[0]
      go_Bordereaux_Add_New(browser)
      print(Client_Short)
      find_client(browser, Client)
      bdx_Cashless_Admission(browser)
  
    bdx_Approved(browser)
    #v = 0
    #for v in range(len(object_list)):
    bdx_Disb_No(browser, object_list[0].Disburment_No)
    RunningNo = object_list[0].Bord_Number

    Policy_temp = Policy.split(" ")
    Policy_short = ""
    for word in Policy_temp:
      Policy_short = Policy_short+str(word[0])

    bdx_bord_No(browser, RunningNo)
    bdx_Date(browser, get_Date())
    if post == True:
      bdx_Input_Id(object_list, browser, adm_base, post = True)
    else:
      bdx_Input_Id(object_list, browser, adm_base)

    bdx_Save(browser)
    Wait(3)
    bdx_Download(browser)
    Wait(10)
    import datetime
    today = datetime.date.today().strftime('%d%m%Y')
    file = RunningNo.split("/")[0] + " " + today
  return file
