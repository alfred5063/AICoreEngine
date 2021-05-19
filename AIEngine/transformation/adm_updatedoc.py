#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - CBA/ADMISSION

from transformation.adm_bdx_updatecm import *
from transformation.adm_bdx_updatedcm import *
from transformation.adm_connect_db import *
from automagic.marc_adm import *
import os
from directory.createfolder import createfolder 
from transformation.adm_bdx_manipulation import get_specific_insurer_detail
from transformation.adm_fill_bdx import get_client_db_variables,get_all_db_invoice_type,get_specific_invoice_type
from utils.logging import logging
import traceback
class adm:
  caseid = 0
  post_main_caseid = 0
  GL_Amount=0
  GL_Amount_post=0
  ID_Status=""
  Status=False
  Post=False
  Date=""
  Double_Billing=False
  First_call=False
  client_name=""
  policy_type=""
  invoice_type=""

  Bord_Number=""
  Disburment_No=""
  att=""
  add=""
  bdx_type=""

  ws_path=""
  First_call_path=""
  Ac_ws=""
  Dr_Fee_Non=""
  Dr_Fee=""
  bdx_path=""
  ack_path=""
  def __init__(self, caseid):
    print("creating object")
    self.caseid=caseid
#database
def get_client_detail_to_excel(browser,adm_base,BDFile):
  df=read_BdxL(BDFile)
  first_case_id=df["File No"][0]
  if "-" in first_case_id:
    first_case_id=first_case_id.split("-")[1]
  object_list=[adm(first_case_id)]
  

  if Check_post_or_type(object_list[0].caseid)=="POST":
    object_list[0].Post=True

  try:
    go_Inpatient_Admission_View(browser)
    audit_log("Marc login login successfuly", "Completed...", adm_base)
  except:
    audit_log("Marc failed to login", "Completed...", adm_base)

  if object_list[0].Post:
    go_Inpatient_Admission_Post_View(browser)
    key_InID(first_case_id,browser)
    click_PaitientName_post(browser)
    Client_Name=get_Client_Name_post(browser)
    pol=get_Pol_Type_post(browser)
  else:
    go_Inpatient_Admission_View(browser)
    key_InID(first_case_id,browser)
    click_PaitientName(browser)
    Client_Name= get_Client_Name(browser)
    pol=get_Pol_Type(browser)

  object_list[0].policy_type=pol
  object_list[0].client_name=Client_Name
  
  Cliend_id=get_specific_insurer_detail(object_list[0],adm_base)[0][4]
  print(object_list[0].Post,"?")
  invoice_obj=get_client_db_variables(browser,object_list,adm_base)
  all_invoice_type=get_all_db_invoice_type(Cliend_id,adm_base)
  bdx_type=get_specific_invoice_type(all_invoice_type,invoice_obj,adm_base)[0]
  print(Cliend_id,bdx_type)
  return bdx_type,Cliend_id


def update_bord_to_database(BDFile, Cliend_ID, invoice_type, adm_base):
  #try:
  #  bdx_to_CM_database(BDFile, invoice_type, Cliend_ID, adm_base)
  #  audit_log("Admission - CM database is updated", "Completed...", adm_base)
  #except:
  #  audit_log("Admission - CM database update is failed", "Completed...", adm_base)
  #  logging("ADM-process admission", traceback.format_exc(), adm_base)
  try:
    bdx_to_DCM_database(BDFile, invoice_type, adm_base)
    audit_log("Admission - DCM database is updated", "Completed...", adm_base)
  except:
    audit_log("Admission - DCM database update is failed", "Completed...", adm_base)
    logging("ADM-process admission", traceback.format_exc(), adm_base)

#adm process 1
def create_path(print_path, taskid, base, email):
  if os.path.exists(print_path) != True:
    createfolder(print_path, base)
  if email == None or email == "":
    task_name_path = print_path + "\\{}".format(get_Task_Name(taskid) + "_" + taskid)
  else:
    email=email.split("@")[0]
    task_name_path=print_path+"\\{}".format(get_Task_Name(taskid)+"_"+email)
  if os.path.exists(task_name_path) != True:
    createfolder(task_name_path,base)

  print_archieve=task_name_path+"\\Archieved"
  print_new=task_name_path+"\\New"

  if os.path.exists(print_archieve) != True:
    createfolder(print_archieve,base)
  if os.path.exists(print_new) != True:
    createfolder(print_new,base)
  return print_new,print_archieve


#adm process 2 get_CMFile_Location(48,"Post")
def update_doc(BDFile,local_CMFile,local_DCMFile,Cliend_ID,bdx_type,adm_base):
  CMFile_Location=get_CMFile_Location(Cliend_ID,bdx_type)[0][0].replace("V:\\",local_CMFile)#V:\ALAM FLORA\ALAM FLORA_MASTER FILE 2019.xlsx
  print(CMFile_Location)
  DCMFile_Location=local_DCMFile
  try:
    if os.path.exists(CMFile_Location):
      audit_log("CM file location is found", "Completed...", adm_base)
      audit_log(CMFile_Location, "Completed...", adm_base)
      bdx_to_CM(BDFile,CMFile_Location,bdx_type,adm_base)#path=CMFile_Location_dummy
    else:
      audit_log("Admission - CM file is not found", "Completed...", adm_base)
  except:
    audit_log("Admission - CM file is failed to insert data", "Completed...", adm_base)

  try:
    DCMFile_Location=local_DCMFile
    if os.path.exists(DCMFile_Location):
      audit_log("Admission - DCMfile location is".format(DCMFile_Location), "Completed...", adm_base)
      bdx_to_DCM(BDFile,DCMFile_Location,2,adm_base)
    else:
      audit_log("Admission - DCM file is not found", "Completed...", adm_base)
  except:
    audit_log("Admission - DCMfile location is failed to insert data", "Completed...", adm_base)


#print listing p3
def insert_print_lisiting(taskid,caseid,file_name,owner):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  cursor.execute('insert into dbo.print_listing(taskid,caseid,filename,created_by,status) values(%s, %s, %s, %s,%s)', (taskid,caseid,file_name,owner,'New'))
  connector.commit()

def update_print_to_db(object_list,owner,taskid):
  for i in object_list:
    if i.caseid==344169:
      continue
    if i.First_call_path is not "":
      insert_print_lisiting(taskid,i.caseid,i.First_call_path+".pdf",owner)
    if i.Ac_ws is not "":
      insert_print_lisiting(taskid,i.caseid,i.Ac_ws+".pdf",owner)
    if i.Dr_Fee_Non is not "":
      insert_print_lisiting(taskid,i.caseid,i.Dr_Fee_Non+".pdf",owner)
    if i.Dr_Fee is not "":
      insert_print_lisiting(taskid,i.Dr_Fee,i.Ac_ws+".pdf",owner)
    if i.ack_path is not "":
      insert_print_lisiting(taskid,i.caseid,i.ack_path+".pdf",owner)
    if i.ack_path_db is not "":
      insert_print_lisiting(taskid,i.caseid,i.ack_path_db+".pdf",owner)


    if i.ws_path is not "":
      insert_print_lisiting(taskid,i.caseid,i.ws_path+".xlsm",owner)
    if i.bdx_path is not "":
      insert_print_lisiting(taskid,i.caseid,i.bdx_path,owner)
    if i.bdx_path_db is not "":
      insert_print_lisiting(taskid,i.caseid,i.bdx_path_db,owner)



