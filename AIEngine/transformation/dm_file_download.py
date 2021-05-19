#!/usr/bin/python
# FINAL SCRIPT updated as of 18th Dec 2020
# Workflow - STP-DM

# Declare Python libraries needed for this script
from utils.Session import session
from directory.createfolder import createfolder
from utils.logging import logging
from utils.audit_trail import audit_log
from directory.get_filename_from_path import get_file_name
from directory.directory_setup import prepare_directories
import os
import zipfile

def prepare_file_folder(DM_base, source=None): #create directory based on user email
  email = DM_base.email
  user_folder = source + '\\' + email
  if(os.path.isdir(user_folder)==False):
    createfolder(user_folder, DM_base, parents=True, exist_ok=True)
    print("Creating directory: New Folder -> %s" % user_folder)
    directory = prepare_directories(user_folder, DM_base)
    return directory
  else:
    directory = prepare_directories(user_folder, DM_base)
    return directory
                             
def detect_datatype(DM_base, filename=None): #detect datatype file
  try:
    if filename.endswith('.zip') or filename.endswith('.RAR') or filename.endswith('.7z'):
      return True
    elif filename.endswith('.xlsx') or filename.endswith(".xls") or filename.endswith('.XLSX') or filename.endswith('.XLS'):
      return False
    else:
      pass
    audit_log("STP-DM - Detect file data type.", "Completed...", DM_base)
  except Exception as error:
    logging("STP-DM - Detect file data type.", error, DM_base)

def unzip_file(source_filename, DM_base,dest_dir, password=None):
  try:
    if password == None:
      with zipfile.ZipFile(source_filename,'r') as zf:
          zf.extractall(str(dest_dir))
    else:
      with zipfile.ZipFile(source_filename,'r') as zf:
          zf.extractall(str(dest_dir), pwd = bytes(password, 'utf-8'))
    return dest_dir
  except Exception as error:
    logging("STP-DM - Unzip file .", error, DM_base)
