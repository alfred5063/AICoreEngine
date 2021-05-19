#!/usr/bin/python
# FINAL SCRIPT updated as of 11th NOVEMEBR 2019
# Workflow - CBA/print listing
# Version 1
from connector.connector import MSSqlConnector
import pandas as pd
from datetime import datetime
import os
import ast
import traceback
import json

def process_get_print_listing_new():
  try:
    connector = MSSqlConnector()
    cursor = connector.cursor()
    qry="select title,caseid,filename,filepath,created_by,created_on from [dbo].[vw_print_listing] where status='New'"
    cursor.execute(qry)
    records = cursor.fetchall()
    df = pd.DataFrame(columns=list(["Case_ID","Task","File_Name","File_Path","Created_by","Created_on"]))
    counter=0
    for each in records:
      #la=each[3].replace("\\\\","\\")
      #print(la)
      df.loc[counter]=[each[1],each[0],each[2],each[3],each[4],each[5].strftime("%m-%d-%Y %H:%M:%S")]
      counter+=1
    x=df.to_json(orient="records")
    json_ob=json.loads(x)
    cursor.close()
    connector.close()
    return json_ob
  except:
    print(traceback.format_exc())
    cursor.close()
    connector.close()
    return traceback.format_exc()

def process_move_files_to_archieved(files_name,email):
  if isinstance(files_name,str):
    print("conberting")
    files_name=ast.literal_eval(files_name)
  print(type(files_name))
  print((files_name))
  connector = MSSqlConnector()
  cursor = connector.cursor()
  now = str(datetime.now().strftime("%m-%d-%Y-%H-%M-%S"))
  for file in files_name:
    cursor.execute("update [dbo].[vw_print_listing] set status = 'Archieved',modified_by = %s,modified_on =GETDATE() where filename=%s",(email,file))
  connector.commit()
  for file in files_name:
    try:
      cursor.execute("select filepath from [dbo].[vw_print_listing]  where filename=%s",(file))
      records = cursor.fetchall()[0][0]
    except:
      continue
    cursor.execute("update [dbo].[vw_print_listing] set filename = %s where filename=%s",(file+"_"+now,file))
    connector.commit()
    new_path=records.replace('New','Archieved')
    if os.path.exists(records+'.pdf'):
      try:
        os.rename(records+'.pdf',new_path+"_"+now+'.pdf')
      except:
        continue
  cursor.close()
  connector.close()
  return {}

def process_get_print_listing_archieved():
  connector = MSSqlConnector()
  cursor = connector.cursor()
  qry="select title,caseid,filename,filepath,created_by,created_on from [dbo].[vw_print_listing] where status='Archieved'"
  cursor.execute(qry)
  records = cursor.fetchall()
  df = pd.DataFrame(columns=list(["Case_ID","Task","File_Name","File_Path","Created_by","Created_on"]))
  counter=0
  for each in records:
    try:
      df.loc[counter]=[each[1],each[0],each[2],each[3],each[4],each[5].strftime("%m-%d-%Y %H:%M:%S")]
    except:
      df.loc[counter]=[each[1],each[0],each[2],each[3],each[4],each[5]]
    counter+=1
  x=df.to_json(orient='records')
  json_ob=json.loads(x)
  cursor.close()
  connector.close()
  return json_ob

