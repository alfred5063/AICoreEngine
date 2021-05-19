# Reading an excel file using Python
import os
import pandas as pd
from directory.get_filename_from_path import get_file_name
from directory.directory_setup import prepare_directories
from utils.audit_trail import audit_log
from utils.logging import logging

#""" Purpose: This function is to use for read Excel file
#    Author: Norhasdayany
#    Created on: 16 May 2019"""
def readExcel(properties, base, columnsname, type='dataframe', **kwargs):
  try:
    audit_log('readExcel', 'import excel as dataframe, filename: %s' % properties.filename, base)
    head, tail = get_file_name(properties.source)
    print("Reading excel file:  %s" % tail)
    prepare_directories(head, base)
    value = ""
    if properties.source.endswith('.xlsx') or properties.source.endswith(".xls") or properties.source.endswith('.XLSX') or properties.source.endswith('.XLS'):
      data = pd.read_excel(properties.source, **kwargs)
      if type == "dataframe":
        value = pd.DataFrame(data, columns= columnsname)
        value = value.astype(str)
    return value
  except Exception as error:
    logging('readExcel', error, base)
    raise error

   
