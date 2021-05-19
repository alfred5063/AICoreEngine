#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - CBA/ADMISSION

# Declare Python libraries needed for this script
from __future__ import unicode_literals
import io
import os
import traceback
import requests
import json
import pandas as pd
from datetime import datetime
from automagic.marc_adm import *
from transformation.adm_cba_status import check_CBA, change_CBA_Status
from transformation.adm_fill_bdx import For_Preview, preview_df, fill_Bdx
from transformation.adm_bdx_manipulation import fileo, bdx_automation, get_all_File, get_Lastest_File
from transformation.adm_lonpac import print_First_call
from transformation.adm_updatedoc import update_doc, create_path, update_print_to_db, update_bord_to_database
from transformation.adm_excel_to_pdf import excel_to_pdf
from pandas.io.json import json_normalize
from utils.Session import session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.logging import logging
from utils.notification import send
from transformation.adm_ws_to_pdf import ws_to_pdf, move_to_storage, pdf_move_to_storage, move_excel_to_print, move_to_bord_listing
from extraction.marc.authenticate_marc_access import get_marc_access
from directory.files_listing_operation import *

myjson = {'parent_guid': '750cfb29-a77f-4ce5-a2c0-55b413995c11', 'guid': '09c48092-1d75-433a-a589-31c392da0111', 'authentication':
          {'username': '', 'password': ''}, 'email': 'alfred.simbun@asia-assistance.com', 'stepname': 'Admission', 'step_parameters': [
  {'id': 'tr_prop_1_step_1', 'key': 'Download_path', 'value': '\\\\dtisvr2\\CBA_UAT\\New', 'description': ''},
  {'id': 'tr_prop_2_step_1', 'key': 'print', 'value': '\\\\dtisvr2\\CBA_UAT\\Print', 'description': ''},
  {'id': 'tr_prop_3_step_1', 'key': 'Result_path', 'value': '\\\\dtisvr2\\CBA_UAT\\Result', 'description': ''},
  {'id': 'tr_prop_4_step_1', 'key': 'Bord_Listing_path', 'value': '\\\\dtisvr2\\CBA_UAT\\Result\\CBA_Paths\\2- BORD LISTING', 'description': ''}],
          'taskId': '360', 'execution_parameters_att_0': '', 'execution_parameters_txt': {'docType': 'ADM', 'parameters': [
            {'Case': '198519'}
            ]}, 'preview': 'False'}
password = str(myjson['authentication']['password'])
username = str(myjson['authentication']['username'])
guid = str(myjson['guid'])
stepname = str(myjson['stepname'])
email = str(myjson['email'])
taskid = myjson['taskId']
preview = str(myjson['preview'])

def workflow_process_admission(myjson, guid, stepname, email, taskid, preview, username, password):
  adm_base = session.base(taskid, guid, email, stepname, username = username, password = password)
  step_data = myjson["step_parameters"]
  download_path = step_data[0]["value"]
  print_path = step_data[1]["value"]
  result_path = step_data[2]["value"]
  Bord_Listing_path = step_data[3]["value"]
  jsondf = pd.DataFrame(json_normalize(myjson['execution_parameters_txt']["parameters"]))
  caseid = list(map(int, jsondf['Case'].tolist()))

  if password != "":
    adm_base.set_password(get_marc_access(adm_base))

  print_new, print_achieve = create_path(print_path, taskid, adm_base, email)
  download_new, download_achieve = create_path(download_path, taskid, adm_base, email)
  audit_log("Admission - Personal folder paths have been created or existed.", "Completed...", adm_base)

  move_to_storage(get_all_File(download_new, "xls"), download_achieve, download_new)
  audit_log("Admission - Files from the users have been moved to a designated path.", "Completed...", adm_base)

  try:
    if preview == 'True':
      try:
        browser = login_Marc(adm_base, download_new)
        object_list = check_CBA(caseid, browser, adm_base)
        audit_log("Admission - Get JSON string from front-end and process Case ID", "Completed...", adm_base)
        invoice_type, object_list, enlisted = For_Preview(browser, object_list, adm_base)
        app_json, empty = preview_df(object_list)
        close_Marc(browser)
        return app_json
      except Exception as error:
        logging("Admission - Failed to provide results for preview. Please check.", error, adm_base)
    else:
      try:
        browser = login_Marc(adm_base, download_new)
        object_list = check_CBA(caseid, browser, adm_base)
        audit_log("Admission - Get JSON string from front-end and process Case ID", "Completed...", adm_base)
        invoice_type, object_list, enlisted = For_Preview(browser, object_list, adm_base)
        app_json, empty = preview_df(object_list)
        if empty == True:
          close_Marc(browser)
          audit_log("Admission - Unable to get client's details.", "Completed...", adm_base)
          return app_json
        else:
          pass_object_list = change_CBA_Status(browser, object_list) ##### Wrongly captured Post as Adm
          audit_log("Admission - Change all Passed Case ID's CBA status to ready", "Completed...", adm_base)
      except Exception as error:
        audit_log("Admission - Failed to process inputs in the backend. Please check.", "Completed...", adm_base)
        logging("Admission - Failed to process inputs in the backend. Please check.", error, adm_base)

      # Generate bordereaux from MARC
      try:
        file_name = fill_Bdx(browser, pass_object_list, adm_base)
      except Exception as error:
        audit_log("Admission - Failed to generate bordereaux from MARC. Please check.", "Completed...", adm_base)
        logging("Admission - Failed to generate bordereaux from MARC. Please check.", error, adm_base)

      time.sleep(20)

      try:
        myfile = list(getListOfFiles(download_new))
        mylist = pd.DataFrame(list(filter(lambda x: 'singBordTemplate.xls' in x, myfile)), columns = ['Filenames'])
        if mylist.empty != True:
          audit_log("Admission - Successfully generated bordereaux from MARC in %s" % download_new, "Completed...", adm_base)

          # Manipulate downloaded bordereaux file
          #Both Adm and Post:
          #If B2B, get Total Bill
          #If not B2B, get Insurance

          try:
            now = str(datetime.now().strftime("%m-%d-%Y-%H-%M-%S"))
            fileobj = fileo()
            fileobj.set_Raw_Path(get_Lastest_File(download_new, "xls"))
            fileobj.set_Ori_Path((download_new+"\{}.xls").format(file_name))
            audit_log("Admission - Successfully manipulated downloaded bordereaux file. Please check.", "Completed...", adm_base)
          except Exception as error:
            audit_log("Admission - Failed to manipulate downloaded bordereaux file. Please check.", "Completed...", adm_base)
            logging("Admission - Failed to manipulate downloaded bordereaux file. Please check.", error, adm_base)

          try:
            status, Cliend_ID = bdx_automation(fileobj, invoice_type, enlisted, adm_base, download_new)
            singbord = download_new + "\\" + "singBordTemplate.xls"
            if os.path.exists(singbord):
              os.remove(singbord)
            else:
              pass
            audit_log("Admission - Successfully removed column and update acknowledgement. Please check.", "Completed...", adm_base)
          except Exception as error:
            audit_log("Admission - Failed to remove column and update acknowledgement. Please check.", "Completed...", adm_base)
            logging("Admission - Failed to remove column and update acknowledgement. Please check.", error, adm_base)

          try:
            BDFile = fileobj.get_Ori_Path()
            move_to_bord_listing(BDFile, Bord_Listing_path, pass_object_list[0], adm_base)
          except Exception as error:
            audit_log("Admission - Failed to move downloaded bordereaux file to designated folder. Please check.", "Completed...", adm_base)
            logging("Admission - Failed to move downloaded bordereaux file to designated folder. Please check.", error, adm_base)


          ## Update bordereaux to database
          try:
            update_bord_to_database(BDFile, Cliend_ID, invoice_type, adm_base)
            # Update to CML is not triggered in this workflow any longer. Need further enhancements as per the request from Si Li.
            os.rename(BDFile, BDFile.replace(".xls", " {}.xls".format(now)))
            BDFile = BDFile.replace(".xls", " {}.xls".format(now))
          except Exception as error:
            audit_log("Admission - Failed to update bordereaux details to database. Please check.", "Completed...", adm_base)
            logging("Admission - Failed to update bordereaux details to database. Please check.", error, adm_base)

        else:
          print("Admission - No generated bordereaux from MARC and download relevent file. Please check.")
          audit_log("Admission - No generated bordereaux from MARC and download relevent file. Please check.", "Completed...", adm_base)
      except Exception as error:
        audit_log("Admission - Failed to generate bordereaux from MARC. Please check.", "Completed...", adm_base)
        logging("Admission - Failed to generate bordereaux from MARC. Please check.", error, adm_base)


      # Send out email to user
      try:
        if email != None:
          send(adm_base, email, "RPA Task for Admission Execution Completed.", "Hi, <br/><br/><br/>You task has been completed. <br/>Reference Number: " + str(guid) + "<br/><br/><br/>Regards,<br/>Robotic Process Automation")
        #adm_base.set_filename(BDFile)
        audit_log(adm_base.stepname, 'Completed', adm_base)
      except Exception as error:
        audit_log("Admission - Failed to send out email to the user. Please check.", "Completed...", adm_base)
        logging("Admission - Failed to send out email to the user. Please check.", error, adm_base)

    close_Marc(browser)

    audit_log("Admission - Process Completed", "Completed...", adm_base)
    return "Completed"

  except Exception as e:
    close_Marc(browser)
    move_to_storage(get_all_File(download_new, "xls"), download_achieve, download_new)
    audit_log("Admission - Failed process. All backend stopped", "Completed...", adm_base)
    logging("Admission - Failed process. All backend stopped.",traceback.format_exc(),adm_base)
    return traceback.format_exc()
