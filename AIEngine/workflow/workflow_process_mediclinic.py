#!/usr/bin/python
# FINAL SCRIPT updated as of 04th NOVEMEBR 2019
# Workflow - CBA/MEDICLINIC
# Version 1
from transformation.medi_file_manipulation import file_manipulation,all_excel_move_to_archive,create_path,createfolder,move_file,get_all_File
from transformation.medi_medi import Medi_Mos
from transformation.medi_bdx_manipulation import bdx_automation
from transformation.medi_update_dcm import update_to_dcm
from transformation.medi_generate_dc import generate_dc
from pandas.io.json import json_normalize
from utils.Session import session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.logging import logging
from extraction.marc.authenticate_marc_access import get_marc_access
import pandas as pd
import traceback,os
from utils.notification import send
from datetime import datetime

from openpyxl import *

#requirement developer mode
disbursementMaster = r"C:\Users\CHUNKIT.LEE\Desktop\test\Disbursement Claims Running No 2019.xls"
#disbursementClaim
disbursementClaim = r"C:\Users\CHUNKIT.LEE\Desktop\test\MCLXXXXX.xls"
# Bordereaux Listing
bordereauxListing = r"C:\Users\CHUNKIT.LEE\Desktop\test\AETNA11324-2019-09 WEB.xls"
def workflow_process_mediclinic(disbursementClaim,bordereauxListing):
    
    # Open the workbook 
    wb = load_workbook(disbursementClaim)
    print(wb)








    for i in range(len(get_all_File(uploadfile,"xls"))):
      try:
        #bdx_type=Medi_Mos(download_new,medi_base,bdx_type,Client,Insurance,input)
        #audit_log("Enter data into Mediclinic website", "Completed...", medi_base)
        file_name,file_path, download_achieve=file_manipulation(download_new,download_new,download_achieve,medi_base,bdx_type)
        audit_log("Extract excel file from zip and rename it", "Completed...", medi_base)
        name,dc_template_path=bdx_automation(file_name,file_path,template_path,medi_base)
        bord_number,ammount,bord_date,running_no,bdx_type=update_to_dcm(file_path,dcm_path)
        audit_log("update data to DCM", "Completed...", medi_base)

        #create new folder for result
        now = datetime.now().strftime('%d%m%Y%H%M%S')
        timestamp = str(now)
        new_result_path = result_path + "\\" + timestamp
        if(os.path.isdir(new_result_path)==False):
          createfolder(new_result_path, medi_base, parents=True, exist_ok=True)

        #generate file and move to result path
        generate_dc(bord_number,ammount,bord_date,running_no,name,bdx_type,dc_template_path,new_result_path,medi_base)
        audit_log("Generate Claim statement", "Completed...", medi_base)
        move_file(file_path,new_result_path+"\\"+file_name)
        audit_log("One File is completed", "Completed...", medi_base)

      except Exception as e:
        audit_log("Failed-All Backend Stopped", "Completed...", medi_base)
        logging("MEDI-process mediclinic error",traceback.format_exc(),medi_base)
        print(traceback.format_exc())
        return traceback.format_exc()
    if email != None:
      send(medi_base, email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. "+
                             "<br/>Reference Number: " + str(guid) +
                             "<br/><br/>Files generated can be found at <"+new_result_path+">"+
                             "<br/>Original file can be found at <"+download_achieve+">"+
                             "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
    audit_log("Process Completed", "Completed", medi_base)
    return "Completed"

def workflow_process_mediclinic(uploadfile):
    for i in range(len(get_all_File(download_new,"zip"))):
      try:
        #bdx_type=Medi_Mos(download_new,medi_base,bdx_type,Client,Insurance,input)
        #audit_log("Enter data into Mediclinic website", "Completed...", medi_base)
        file_name,file_path, download_achieve=file_manipulation(download_new,download_new,download_achieve,medi_base,bdx_type)
        audit_log("Extract excel file from zip and rename it", "Completed...", medi_base)
        name,dc_template_path=bdx_automation(file_name,file_path,template_path,medi_base)
        bord_number,ammount,bord_date,running_no,bdx_type=update_to_dcm(file_path,dcm_path)
        audit_log("update data to DCM", "Completed...", medi_base)

        #create new folder for result
        now = datetime.now().strftime('%d%m%Y%H%M%S')
        timestamp = str(now)
        new_result_path = result_path + "\\" + timestamp
        if(os.path.isdir(new_result_path)==False):
          createfolder(new_result_path, medi_base, parents=True, exist_ok=True)

        #generate file and move to result path
        generate_dc(bord_number,ammount,bord_date,running_no,name,bdx_type,dc_template_path,new_result_path,medi_base)
        audit_log("Generate Claim statement", "Completed...", medi_base)
        move_file(file_path,new_result_path+"\\"+file_name)
        audit_log("One File is completed", "Completed...", medi_base)

      except Exception as e:
        audit_log("Failed-All Backend Stopped", "Completed...", medi_base)
        logging("MEDI-process mediclinic error",traceback.format_exc(),medi_base)
        print(traceback.format_exc())
        return traceback.format_exc()
    if email != None:
      send(medi_base, email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. "+
                             "<br/>Reference Number: " + str(guid) +
                             "<br/><br/>Files generated can be found at <"+new_result_path+">"+
                             "<br/>Original file can be found at <"+download_achieve+">"+
                             "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
    audit_log("Process Completed", "Completed", medi_base)
    return "Completed"


def workflow_process_mediclinic(myjson, guid, stepname, email, taskid):
  data=myjson['execution_parameters_txt']["parameters"][0]
  step_data=myjson["step_parameters"]
  template_path=step_data[0]["value"]
  download_path=step_data[1]["value"]
  excel=step_data[2]["value"]
  dcm_path=step_data[3]["value"]
  bdx_type=data["Type"]

  print(data)
  medi_base = session.base(taskid, guid, email, stepname, username="", password="")
  #download_new,download_achieve=create_path(download_path,medi_base)
  #result_path=download_new.replace("New","Result")
  #result_path=result_path.replace("Result","New",1)

  #if os.path.exists(result_path) != True:
    #createfolder(result_path,medi_base)
  #all_excel_move_to_archive(result_path)
  #list_zip_file=get_all_File(excel,"xls")

  #processing path
  processing_folder = download_path+email+"\\"
  print('source processing: {0}'.format(processing_folder))
  download_new = processing_folder+"New"
  download_achieve = processing_folder+"Archieved"
  result_path = processing_folder+"Result"
  

  #create directory if not exist
  audit_log("Creating directory if not exist.", "Completed...", medi_base)
  if(os.path.isdir(download_new)==False):
      createfolder(download_new, medi_base, parents=True, exist_ok=True)
  if(os.path.isdir(download_achieve)==False):
      createfolder(download_achieve, medi_base, parents=True, exist_ok=True)
  if(os.path.isdir(result_path)==False):
      createfolder(result_path, medi_base, parents=True, exist_ok=True)

  """
  download>extract> automate bdx using insurer name(database)> use file name to update DCM and claim listing
  """
  audit_log("Setup environment successfully", "Completed...", medi_base)
  for i in range(len(get_all_File(download_new,"zip"))):
    try:
      #bdx_type=Medi_Mos(download_new,medi_base,bdx_type,Client,Insurance,input)
      #audit_log("Enter data into Mediclinic website", "Completed...", medi_base)
      file_name,file_path, download_achieve=file_manipulation(download_new,download_new,download_achieve,medi_base,bdx_type)
      audit_log("Extract excel file from zip and rename it", "Completed...", medi_base)
      name,dc_template_path=bdx_automation(file_name,file_path,template_path,medi_base)
      bord_number,ammount,bord_date,running_no,bdx_type=update_to_dcm(file_path,dcm_path)
      audit_log("update data to DCM", "Completed...", medi_base)

      #create new folder for result
      now = datetime.now().strftime('%d%m%Y%H%M%S')
      timestamp = str(now)
      new_result_path = result_path + "\\" + timestamp
      if(os.path.isdir(new_result_path)==False):
        createfolder(new_result_path, medi_base, parents=True, exist_ok=True)

      #generate file and move to result path
      generate_dc(bord_number,ammount,bord_date,running_no,name,bdx_type,dc_template_path,new_result_path,medi_base)
      audit_log("Generate Claim statement", "Completed...", medi_base)
      move_file(file_path,new_result_path+"\\"+file_name)
      audit_log("One File is completed", "Completed...", medi_base)

    except Exception as e:
      audit_log("Failed-All Backend Stopped", "Completed...", medi_base)
      logging("MEDI-process mediclinic error",traceback.format_exc(),medi_base)
      print(traceback.format_exc())
      return traceback.format_exc()
  if email != None:
    send(medi_base, email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. "+
                           "<br/>Reference Number: " + str(guid) +
                           "<br/><br/>Files generated can be found at <"+new_result_path+">"+
                           "<br/>Original file can be found at <"+download_achieve+">"+
                           "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
  audit_log("Process Completed", "Completed", medi_base)
  return "Completed"


