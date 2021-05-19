#!/usr/bin/python
# FINAL SCRIPT to check if excel file can be found and read or not.
# Last updated as of 1st April 2020

# Declare Python libraries needed for this script
import os
import time
from automagica import *
import openpyxl
import datetime
import pandas as pd
from openpyxl import Workbook

# Check if file is locked / used
def is_locked(filepath):
    locked = None
    file_object = None
    if os.path.exists(filepath):
        try:
            buffer_size = 8
            # Opening file in append mode and read the first 8 characters.
            file_object = open(filepath, 'a', buffer_size)
            if file_object:
                locked = False
        except IOError as message:
            locked = True
        finally:
            if file_object:
                file_object.close()
    else:
        print("- FILE: %s not found." % filepath)
    return locked

# Wait for file to open / becomes available
def wait_for_files(filepath):
  wait_time = 600
  for filepath in filepaths:
    while not os.path.exists(filepath):
      time.sleep(wait_time)
    while is_locked(filepath):
      time.sleep(wait_time)

# Force close the application from Task Manager
def force_close_excel(filepath):
  try:
    curtask = pd.DataFrame(os.popen('tasklist').readlines()).to_json(orient = 'index')
    if "excel.exe" in curtask:
      os.system("taskkill /f /im  excel.exe")
    else:
      pass
  except IOError as message:
    print("- File is locked (unable to open in append mode). %s." % message)

