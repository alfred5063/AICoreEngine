#!/usr/bin/python
# FINAL SCRIPT updated as of 18th Dec 2020
# Workflow - STP-DM

# Declare Python libraries needed for this script
import os,time
import datetime
import shutil
import datetime as dt
from directory.deletefile import *
from utils.logging import logging
from utils.audit_trail import audit_log
import pandas as pd
import xlrd
from transformation.dm_password_decrypt import *
from transformation.dm_file_download import *
import glob
import csv
from directory.get_filename_from_path import get_file_name
from xlsxwriter.workbook import Workbook

def read_clientfile(DM_base, filepath):
  file_type = detect_datatype(DM_base, filepath)
  head, tail = get_file_name(filepath)
  try:
    if file_type == False:
      if 'PIB_TPA' in tail.upper():#PACIFIC
        try :
          data = pd.read_excel(filepath)
          value = pd.DataFrame(data)
        except:
          allpass = ['Pib@2019','Pib@2020']
          value = decrypt_password(DM_base, filepath, allpass, file_type)
      elif 'AAN-PIMD' in tail.upper() or 'AAN-PDEL' in tail.upper() or 'AAN-PNEW' in tail.upper(): #LONPAC
        try :
          data = pd.read_excel(filepath)
          value = pd.DataFrame(data)
        except:
          allpass = ['Pib@2019','Pib@2020']
          value = decrypt_password(DM_base, filepath, allpass, file_type)
    elif file_type == True:
      if 'PACIFIC' in head.upper(): #PACIFIC
        filepath_new = unzip_file(filepath, DM_base, head)
        if str('password required for extraction') in str(filepath_new):
          password = ['Pib@2019','Pib@2020']
          for i in range(len(password)):
            filepath_new = unzip_file(filepath, DM_base, head, password[i])
        else:
          pass
      elif 'LONPAC' in head.upper(): #LONPAC
        filepath_new = unzip_file(filepath,DM_base, head)
    audit_log("STP-DM - Read client file." % head, "Completed...", DM_base)
  except Exception as error:
    logging("STP-DM - Read client file.", error, DM_base)

#move file after specific time
def move_file_to_archive(created, dest):
  now = dt.datetime.now()
  ago = now-dt.timedelta(days=2) #example if file > 2 days
  strftime = "%H:%M %m/%d/%Y"
  for root, dirs,files in os.walk(created):
      for fname in files:
          path = os.path.join(root, fname)
          st = os.stat(path)
          mtime = dt.datetime.fromtimestamp(st.st_mtime)
          if mtime < ago:
              print("True:  ", fname, " at ", mtime.strftime("%H:%M %m/%d/%Y"))
              shutil.move(path, dest)
              return fname

#delete file after specific time
def del_file_from_archive(dest, base):
  now = dt.datetime.now()
  ago = now-dt.timedelta(days=2) 
  strftime = "%H:%M %m/%d/%Y"
  for root, dirs,files in os.walk(dest):
      for fname in files:
          path = os.path.join(root, fname)
          st = os.stat(path)
          mtime = dt.datetime.fromtimestamp(st.st_mtime)
          if mtime < ago:
              print("True:  ", fname, " at ", mtime.strftime("%H:%M %m/%d/%Y"))
              deletefile(path, base)
              return fname

def convert_to_xlsx(csvfile,base):
  try:
    head,tail = get_file_name(csvfile)
    sheetname = tail.split(".", 1)[0]
    workbook = Workbook(csvfile[:-4] + '.xlsx')    
    worksheet = workbook.add_worksheet(sheetname)
    with open(csvfile, 'rt', encoding='utf8') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
    workbook.close()
    os.remove(csvfile)
    audit_log("STP-DM - Convert .csv to .xlsx format.", "Completed...", base)
  except Exception as error:
    print(error)
    logging("STP-DM - Error in convert .csv to .xlsx format.", error, base)

