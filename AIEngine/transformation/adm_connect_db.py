#!/usr/bin/python
# FINAL SCRIPT updated as of 22nd June 2020
# Workflow - CBA/ADMISSION
# Version 1

from connector.connector import MSSqlConnector,MySqlConnector
import json
import os
from datetime import date,datetime,timedelta
import pandas as pd

def get_EventDate():
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry="SELECT date FROM cba.calendar_events"
  cursor = connector.cursor()
  cursor.execute(qry)
  records = cursor.fetchall()
  cursor.close()
  connector.close()
  return records


def get_Task_Name(taskid):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry = "SELECT title FROM dbo.tasks where id = {}".format(str(taskid))
  cursor = connector.cursor()
  cursor.execute(qry)
  taskname = cursor.fetchall()
  cursor.close()
  connector.close()
  return taskname[0][0]

def get_lastest_DirNo():
  connector = MSSqlConnector()
  cursor = connector.cursor()
  now = datetime.now().year
  qry="SELECT content FROM cba.disbursement_master where year={}".format(str(now))
  cursor.execute(qry)
  DirNo = cursor.fetchall()
  cursor.close()
  connector.close()
  return DirNo

def get_lastest_BordNum(id,invoice_type):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  now = datetime.now().year
  qry="SELECT content FROM cba.client_master where client_id={} and invoice_type='{}' and year={}".format(id,str(invoice_type),now)
  cursor.execute(qry)
  DirNo = cursor.fetchall()
  cursor.close()
  connector.close()
  return DirNo



def get_text_excluding_children(driver, element):
  ##get element text while the text is not part of any attribute
    return driver.execute_script("""
    return jQuery(arguments[0]).contents().filter(function() {
        return this.nodeType == Node.TEXT_NODE;
    }).text();
    """, element)

def get_insurer_info():
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry="SELECT comp_name,insurer,atten_1,comp_address,id,identifier FROM cba.insurer_details"
  cursor = connector.cursor()
  cursor.execute(qry)
  insurerinfo = cursor.fetchall()
  cursor.close()
  connector.close()
  return insurerinfo

def get_Colomn_Action_Type(id):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  #qry="SELECT action FROM cba.client_column_mapping where client_id="+str(id)
  qry="SELECT action_type FROM cba.client_column_mapping_new where client_id="+str(id)
  cursor = connector.cursor()
  cursor.execute(qry)
  Action_Type = cursor.fetchall()
  cursor.close()
  connector.close()
  return Action_Type[0][0]

def get_Colomn_Involved(id):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  #qry="SELECT action FROM cba.client_column_mapping where client_id="+str(id)
  qry="SELECT action FROM cba.client_column_mapping_new where client_id="+str(id)
  cursor = connector.cursor()
  cursor.execute(qry)
  Colomn_Involved = cursor.fetchall()
  cursor.close()
  connector.close()
  return Colomn_Involved[0][0]

def get_lastest_BordNo(id):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry="SELECT invoice_type,content FROM cba.client_master where client_id="+str(id)
  cursor = connector.cursor()
  cursor.execute(qry)
  DirNo = cursor.fetchall()
  cursor.close()
  connector.close()
  return DirNo

def get_CMFile_Location(id,bdx_type):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry="SELECT url FROM cba.client_master where client_id=%s and invoice_type=%s"
  cursor = connector.cursor()
  cursor.execute(qry,(id,bdx_type))
  CMFile_Location = cursor.fetchall()
  cursor.close()
  connector.close()
  return CMFile_Location

def get_DCMFile_Location():
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry="SELECT url FROM cba.disbursement_master"
  cursor = connector.cursor()
  cursor.execute(qry)
  DCMFile_Location = cursor.fetchall()
  cursor.close()
  connector.close()
  return DCMFile_Location

def post_get_main_case_id(caseid):
  conn = MySqlConnector()
  cur = conn.cursor()

  # Calling Stored Procedure
  paramater = [str(caseid),]
  stored_proc = cur.callproc('adm_query_marc_by_caseid', paramater)
  for i in cur.stored_results():
    results = i.fetchall()
  fetched_result = pd.DataFrame(results)

  cur.close()
  conn.close()

  if len(fetched_result) != 0:
    return fetched_result[0][0]
  else:
    return fetched_result


def print_file(path):
    print("FILE LIST")
    files = []
    # r=root, d=directories, f = files
    for r, d, f in os.walk(path):
        for file in f:
            if '.pdf' in file:
                files.append(os.path.join(r, file))
    PREVIEW_FILE=[]
    for f in files:
        PREVIEW_FILE.append(f.replace(path+"\\",""))

    for f in PREVIEW_FILE:
      print(f)

def update_dcm_to_database(json_dcm,email):
  current_year = int(datetime.now().year)
  print(current_year)
  connector = MSSqlConnector()
  cursor = connector.cursor()
  cursor.execute("select content from [cba].[disbursement_master] where [year] = %s",(current_year))
  dictionary_db=json.loads(cursor.fetchall()[0][0])
  list_db=dictionary_db["Data"]
  list_db.append(json_dcm.copy())
  result=json.dumps(dictionary_db)
  cursor.execute("update [cba].[disbursement_master] set content = %s , modifiedby=%s where [year] = %s",(result,email,current_year))
  connector.commit()
  cursor.close()
  connector.close()
  return result

def update_new_dcm_to_database(json_dcm,email):
  current_year = int(datetime.now().year)
  last_year=current_year-1
  date=(datetime.now().date())
  connector = MSSqlConnector()
  cursor = connector.cursor()
  cursor.execute("select url from [cba].[disbursement_master] where [year] = %s",(last_year))
  url=cursor.fetchall()[0][0]
  url=url.replace(str(last_year),str(current_year))
  dictionary_db={"Data": []}
  list_db=dictionary_db["Data"]
  list_db.append(json_dcm.copy())
  result=json.dumps(dictionary_db)
  cursor.execute("""INSERT INTO [cba].[disbursement_master] (content, year, createdby,createdon,url)
                  VALUES (%s,%s,%s,%s,%s)"""
                 ,(result,str(current_year),email,date,url))
  connector.commit()
  cursor.close()
  connector.close()
  return result


def update_cm_to_database(data,cliend_id,type1,email):
  connector = MSSqlConnector()
  cursor = connector.cursor()
  current_year= datetime.now().year
  cursor.execute("select content from [cba].[client_master] where [client_id] = %s and [invoice_type]=%s and year=%s",(cliend_id,type1,current_year))
  dictionary_db=json.loads(cursor.fetchall()[0][0])
  list_db=dictionary_db["Data"]
  list_db.append(data)
  result=json.dumps(dictionary_db)
  cursor.execute("update [cba].[client_master] set content = %s,modifiedby=%s where [client_id] = %s and [invoice_type]=%s",(result,email,cliend_id,type1))
  connector.commit()
  cursor.close()
  connector.close()
  return result

def update_new_cm_to_database(data,cliend_id,type1,email):
  current_year = int(datetime.now().year)
  date=(datetime.now().date())
  last_year=current_year-1
  connector = MSSqlConnector()
  cursor = connector.cursor()
  cursor.execute("select url from [cba].[client_master] where [client_id] = %s and [invoice_type]=%s",(cliend_id,type1))
  url=cursor.fetchall()[0][0]
  url=url.replace(str(last_year),str(current_year))
  dictionary_db={"Data": []}
  list_db=dictionary_db["Data"]
  list_db.append(data.copy())
  result=json.dumps(dictionary_db)
  cursor.execute("""INSERT INTO [cba].[client_master] (year,client_id,invoice_type,content, createdby,createdon,url)
                  VALUES (%s,%s,%s,%s,%s,%s,%s)"""
                 ,(str(current_year),cliend_id,type1,result,email,date,url))
  connector.commit()
  cursor.close()
  connector.close()
  return result
