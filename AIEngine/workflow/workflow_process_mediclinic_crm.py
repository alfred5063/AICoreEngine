#!/usr/bin/python
# FINAL SCRIPT updated as of 04th NOVEMEBR 2019
# Workflow - CBA/MEDICLINIC
# Version 1
from transformation.medi_file_manipulation import file_manipulation_crm,all_excel_move_to_archive,create_path,move_file
from transformation.medi_crm_medi import Medi_Crm
from transformation.medi_crm_bdx_manipulation import bdx_automation
from transformation.medi_update_dcm import update_to_dcm
from transformation.medi_generate_dc import generate_dc
from pandas.io.json import json_normalize
from utils.Session import session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.logging import logging
from extraction.marc.authenticate_marc_access import get_marc_access
import pandas as pd
import traceback
from utils.notification import send
from directory.createfolder import createfolder
import os

def workflow_process_mediclinic_crm(myjson, guid, stepname, email, taskid,username,password, debug=False):
  
  medi_base = session.base(taskid, guid, email, stepname, username=username, password=password)
  audit_log("Initial the mediclinic crm rpa process.", "Completed...", medi_base)
  data=myjson['execution_parameters_txt']["parameters"][0]
  data_caseid=myjson['execution_parameters_txt']["parameters"]
  step_data=myjson["step_parameters"]
  #client, download=new, result, archived, disbursement claim master path
  #previous parameters for path creation
  #template_path=step_data[0]["value"]
  #download_path=step_data[1]["value"]
  #excel=step_data[2]["value"]
  #dcm_path=step_data[3]["value"]
  for item in step_data:
      if item['key']=='source':
        source = str(item['value'])
      elif item['key']=='dcm_path':
        dcm_path = str(item['value'])

  #processing path
  processing_folder = source+email+"\\"
  print('source processing: {0}'.format(processing_folder))
  new_folder = processing_folder+"New"
  archieved_folder = processing_folder+"Archieved"
  result_folder = processing_folder+"Result"
  client_folder = source+"\\Client"

  #create directory if not exist
  audit_log("Creating directory if not exist.", "Completed...", medi_base)
  if(os.path.isdir(new_folder)==False):
      createfolder(new_folder, medi_base, parents=True, exist_ok=True)
  if(os.path.isdir(archieved_folder)==False):
      createfolder(archieved_folder, medi_base, parents=True, exist_ok=True)
  if(os.path.isdir(result_folder)==False):
      createfolder(result_folder, medi_base, parents=True, exist_ok=True)
  bdx_type=data["Type"]
  option=data["Option"]
  if bdx_type.lower()=="manual":
    Client=data["Client"]
    if option.lower() =="insurance":
      Insurance=data["Insurance"]
    else:
      Insurance=None
    input=list(i['Case ID'] for i in data_caseid)
  elif bdx_type.lower()=="web":
    Client=data["Client"]
    if option.lower() =="insurance":
      Insurance=data["Insurance"]
    else:
      Insurance=None
    Start_date=data["Start Date"]
    End_date=data["End Date"]
    input=[Start_date,End_date]

  #download_new,download_achieve=create_path(download_path,medi_base)
  audit_log("Predefined variables to process next step.", "Completed...", medi_base)
  #audit_log(download_achieve, "Completed...", medi_base)
  #all_excel_move_to_archive(excel)
  if debug==False:
    if password != "":
      medi_base.set_password(get_marc_access(medi_base))
      #audit_log("Password is decryted", "Completed...", medi_base)
      #medi_base.set_password(password)
  """
  download>extract> automate bdx using insurer name(database)> use file name to update DCM and claim listing
  """
  audit_log("Open to CRM site.", "Completed...", medi_base)
  print('Open CRM site.')
  try:
    Medi_Crm(new_folder,medi_base,bdx_type,Client,Insurance,input,option)
    #audit_log("Download path is ".format(download_new), "Completed", medi_base)

    file_name,file_path=file_manipulation_crm(new_folder,archieved_folder,medi_base)

    dc_template_path,name,file_path=bdx_automation(file_name,file_path,template_path,bdx_type,medi_base)
    if file_path==False:
      audit_log("The excel downloaded do not have any content. Completed RPA transaction.", "Completed...", medi_base)
      return "Completed"
    audit_log("The unused excel column is removed", "Completed...", medi_base)

    bord_number,ammount,bord_date,running_no,bdx_type=update_to_dcm(file_path,dcm_path,ocm=True)
    audit_log("updated data to DCM", "Completed...", medi_base)

    generate_dc(bord_number,ammount,bord_date,running_no,name,bdx_type,dc_template_path,excel)
    audit_log("Generate Claim statement", "Completed...", medi_base)

    move_file(file_path,file_path.replace(download_new,excel))
    audit_log("Process Completed", "Completed...", medi_base)
    if email != None:
      send(medi_base, email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
    audit_log("Process Completed", "Completed", medi_base)
    return "Completed"
  except Exception as e:
    audit_log("Failed-All Backend Stopped", "Completed...", medi_base)
    logging("MEDI-process mediclinic error",traceback.format_exc(),medi_base)
    return traceback.format_exc()
