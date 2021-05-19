import pandas as pd
import numpy as np
from xlrd import open_workbook
import xlrd
import os
from directory.movefile import copy_file_to_archived, move_file_to_result_medi
from directory.get_filename_from_path import get_file_name
from utils.Session import session
from utils.audit_trail import *
from utils.logging import logging
from utils.notification import send
from directory.directory_setup import prepare_directories
from extraction.marc.mediclinic_kpi_query import get_claim_infor, get_staff_productivity
from directory.get_filename_from_path import get_file_name
from directory.directory_setup import prepare_directories
from utils.notification import send
from loading.excel.createexcel import create_excel
import time

#taskid = 1
#guid = 2
#step_name = 'test'
#start_approval_date = '2021-01-01'
#end_approval_date = '2021-01-31'
#case_created_date = '2021-01-01'
#claim_created_date = '2021-01-01'
#base = '1'
#source = r'C:\Users\CHOONLAM.CHIN\Desktop\tes,,,,\mediclinic'
#email = 'choonlam.chin@asia-assistance.com'
#destination = r'C:\Users\CHOONLAM.CHIN\Desktop\tes,,,,\mediclinic'


def generate_report(taskid, guid, step_name, claim_status, start_visit_date, end_visit_date,claim_type,
                          case_type, start_approval_date, end_approval_date,source,email,destination=None):
  base = session.base(taskid, guid,email,step_name)
  head, tail = get_file_name(source)
  #to validate the directory whether have new, result and archive. If not, then create a new directory
  prepare_directories(head+"\\"+tail, base)
  try:
    #retrieve query from mysql -> preproduction environment - separate connection
    df = get_claim_infor(base, claim_status=claim_status, start_visit_date=start_visit_date, end_visit_date=end_visit_date,
                         claim_type=claim_type, case_type=case_type, start_approval_date=start_approval_date, end_approval_date=end_approval_date)
    if not df.empty:

      df.insert(0, 'ID', range(1, 1 + len(df)))
      df.rename(columns={'Case_Id':'Case Id','Claim_Id':'Claim Id',	'Mm_Name': 'Mm Name',	'Mm_National_Id':'Mm National Id',
                         'Mm_Other_National_Id':'Mm Other National Id',	'Company_Or_Client_Name': 'Company/Client Name','parent_company':'Parent Company',
                         'Insurer_Name':'Insurer Name',	'Claim_Type':'Claim Type',	'Case_Type':'Case Type',
                         'Claim_Status': 'Claim Status', 	'Visit_Date':'Visit Date',	'Clinic_Or_Hospital':'Clinic/Hospital',
                         'Other_Clinic':'Other Clinic',	'Total_Bill_Amt':'Total Bill Amt',	'Insurance_Amt':'Insurance Amt',
                         'Patient_Amt':'Patient Amt',	'ASO_Amt':'ASO Amt',	'Created_Date':'Created Date',
                         'Bill_Received_Date':'Bill Received Date',	'Prev_Statement_No':'Prev Statement No',
                         'Approval_Type':'Approval Type',	'Approval_Date': 'Approval Date',	'Approval_Amount':'Approval Amt',
                         'Created_By':'Created By',	'Last_Edited_By':'Last Edited By',	'outptclaimapprovalbyabbvname':'Claim Approval Abbv Name'}, inplace=True)
      source = head+"\\"+tail
      destination = head+"\\"+tail+"\\Result\\"
      print("file path source: {0}; destination: {1}".format(source, destination))
      properties = session.data_management(source, destination)
  
      timestr = time.strftime("%Y%m%d-%H%M%S")
      filename = "result_claim_utilization_"+str(timestr)+".xlsx"
      #assign full path that later use to generate excel file with timestamp
      fullpath = head+"\\"+tail+"\\Result\\"+ filename
      #print(fullpath)

      #create new excel file with query that retrieved in data frame format
      create_excel(properties, df, base, sheetname ='Claim Utilization', filename=fullpath)

      #send notification email with path to download
      send(base, base.email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>Your task has been completed. "+
                          "<br/>Reference Number: " + str(guid) + "<br/>"+
                          "<br/>Report can be download from the following link: <"+fullpath+"><br/>"+
                          "<br/><br/>"+
                          "<br/>Regards,<br/>Robotic Process Automation", file = fullpath)
    else:
      send(base, base.email, "RPA Task Execution Completed.","Hi, <br/><br/><br/>Your task has been completed. "+
                          "<br/>Reference Number: " + str(guid) + "<br/>"+
                          "<br/>Your filtered result return empty. <br/>"+
                          "<br/><br/>"+
                          "<br/>Regards,<br/>Robotic Process Automation")
  except Exception as error:
    print('Error: {0}'.format(error))
    return 'Fail'
  return 'Success'

def generate_staff_productivity_report(taskid, guid, step_name, start_approval_date, end_approval_date,source,email,destination=None):
  base = session.base(taskid, guid,email,step_name)
  head, tail = get_file_name(source)
  #to validate the directory whether have new, result and archive. If not, then create a new directory
  prepare_directories(head+"\\"+tail, base)
  try:
    #retrieve query from mysql -> preproduction environment - separate connection
    df, spgl_df = get_staff_productivity(base, start_approval_date=start_approval_date, end_approval_date=end_approval_date)
    pd.set_option('max_columns', None)
    #result_df.to_excel(r'C:\Users\NURFAZREEN.RAMLY\Desktop\RPA-TEST\test3.xlsx')
    new_df = []
    if not df.empty:
      rows = df.to_dict(orient="records")

      for row in rows: 
        if row['Case Type'] == 'Cashless':
          if row['Claim Type']=='Outpt - GP':
            row['Cashless GP'] = 1
          else:
            if row['Approval Type'] == 'A' or row['Approval Type'] == 'R':
              row['Cashless SP'] = 1
            else:
              pass
        elif row['Case Type'] == 'Cashless Chronic' or row['Case Type'] == 'Cashless - Special Arrangement':
          if row['Claim Type']=='Outpt - GP':
            row['Cashless GP'] = 1
          else:
            if row['Approval Type'] == 'A' or row['Approval Type'] == 'R':
              row['Cashless SP'] = 1
            else:
              pass
        elif row['Case Type']=='Reimbursement' or row['Case Type']=='Reimbursement - Chronic' or row['Case Type']=='Reimbursement - Special Arrangement':
          row['Reimbursement'] = 1
        elif row['Case Type']=='External':
          row['External'] = 1

        #if row['outptClaimOutptStatusId'] == 1:
        #    if row['outptcasedocname'] == 'SPGL Initial' or row['outptcasedocname'] == 'SPGL Final':
        #      row['SPGL'] = 1

        new_df.append(row)

      #if not spgl_df.empty:
      #  rows = spgl_df.to_dict(orient="records")
      #  for row in rows:
      #    if row['outptClaimOutptBenefitName'] == 'SP' and row['outptcasedocname'] == 'SPGL Initial' or row['outptcasedocname'] == 'SPGL Final':
      #      row['SPGL'] = 1

          #new_df.append(row)

      result = pd.DataFrame(new_df)
      #fill in empty to 0
      result.fillna(0, inplace=True)
      result.drop(columns=['outptClaimOutptStatusId'], inplace=True)
      result=result.groupby(['Approval Date','Staff Name']).sum().reset_index()

      source = head+"\\"+tail
      destination = head+"\\"+tail+"\\Result\\"
      print("file path source: {0}; destination: {1}".format(source, destination))
      properties = session.data_management(source, destination)
  
      timestr = time.strftime("%Y%m%d-%H%M%S")
      filename = "result_staff_productivity_"+str(timestr)+".xlsx"
      #assign full path that later use to generate excel file with timestamp
      fullpath = head+"\\"+tail+"\\Result\\"+ filename
      #print(fullpath)

      #create new excel file with query that retrieved in data frame format
      create_excel(properties, result, base, sheetname ='Staff Productivity', filename=fullpath)
      #base.email = email
      #send notification email with path to download
      send(base, base.email, "RPA Task Execution Completed.", "Hi, <br/><br/><br/>Your task has been completed. "+
                          "<br/>Reference Number: " + str(guid) + "<br/>"+
                          "<br/>Report can be download from the following link: <"+fullpath+"><br/>"+
                          "<br/><br/>"+
                          "<br/>Regards,<br/>Robotic Process Automation", file = fullpath)
    else:
      send(base, base.email, "RPA Task Execution Completed.","Hi, <br/><br/><br/>Your task has been completed. "+
                          "<br/>Reference Number: " + str(guid) + "<br/>"+
                          "<br/>Your filtered result return empty. <br/>"+
                          "<br/><br/>"+
                          "<br/>Regards,<br/>Robotic Process Automation")
  except Exception as error:
    print('Error: {0}'.format(error))
    return 'Fail'
  return 'Success'
