#!/usr/bin/python
# FINAL SCRIPT updated as of 13th January 2021

# Declare Python libraries needed for this script
from directory.createfolder import createfolder
import os
from utils.logging import logging
from utils.audit_trail import audit_log

def prepare_directories(path, base):
  try:
    print("Creating directory...")
    New = ""
    Archived = ""
    Result = ""
   
    if("New" in path) or ("new" in path):
      New = path
      before = path.rfind('\\')
      path = path[:before]
      Archived = path + '\\Archived'
      Result = path + '\\Result'
    else: 
      New = path + '\\New'
      Archived = path + '\\Archived'
      Result = path + '\\Result'
   
    if(os.path.isdir(New)==False):
      createfolder(New, base, parents=True, exist_ok=True)
      print("Creating directory: New Folder -> %s" % New)
   
    if(os.path.isdir(Archived)==False):
      createfolder(Archived, base, parents=True, exist_ok=True)
      print("Creating directory: Archived Folder -> %s" % Archived)

    if(os.path.isdir(Result)==False):
      createfolder(Result, base, parents=True, exist_ok=True)
      print("Creating directory: Result Folder -> %s" % Result)
   
  except Exception as error:
    logging('Creating Directories', error, base)
    print('An error occured while creating directories: ', error)

def prepare_directories_dm(path, currentdate, timestamp, base):
  try:
    print("Creating directory...")
    MAIN = path + '\\RUN_' + currentdate
    CURRENTDATE = ""
    FAILED = ""
    ARCHIVED = ""
    RESULTS = ""
   
    if(str(currentdate) in path):
      FAILED = path + '\\RUN_' + currentdate + '\\FAILED'
      ARCHIVED = path + '\\RUN_' + currentdate +  '\\ARCHIVED'
      RESULTS = path + '\\RUN_' + currentdate +  '\\RESULTS'
    else:
      CURRENTDATE = path + '\\RUN_' + currentdate
      FAILED = path + '\\RUN_' + currentdate + '\\FAILED'
      ARCHIVED = path + '\\RUN_' + currentdate +  '\\ARCHIVED'
      RESULTS = path + '\\RUN_' + currentdate +  '\\RESULTS'
   
    if(os.path.isdir(FAILED)==False):
      createfolder(FAILED, base, parents=True, exist_ok=True)

    createfolder(FAILED + "\\FAIL_" + timestamp, base, parents=True, exist_ok=True)
    FAILED = FAILED + "\\FAIL_" + timestamp
   
    if(os.path.isdir(ARCHIVED)==False):
      createfolder(ARCHIVED, base, parents=True, exist_ok=True)

    createfolder(ARCHIVED + "\\ARC_" + timestamp, base, parents=True, exist_ok=True)
    ARCHIVED = ARCHIVED + "\\ARC_" + timestamp

    if(os.path.isdir(RESULTS)==False):
      createfolder(RESULTS, base, parents=True, exist_ok=True)

    createfolder(RESULTS + "\\RES_" + timestamp, base, parents=True, exist_ok=True)
    RESULTS = RESULTS + "\\RES_" + timestamp
   
  except Exception as error:
    logging('Creating Directories', error, base)
    print('An error occured while creating directories: ', error)

  return FAILED, ARCHIVED, RESULTS
