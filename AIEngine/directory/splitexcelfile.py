import pandas as pd
from utils.audit_trail import audit_log
from utils.logging import logging
from extraction.excel.readexcel import readExcel
from directory.get_filename_from_path import get_file_name
import os
from pandas import ExcelWriter
from openpyxl import Workbook

def split_excel_file(source, destination, row_limit=250, keep_headers=True):
    files_to_process = []
    source_file = source
    #print('filepath: {0}', source_file)
    path, filename = get_file_name(source_file)
    base_filename = os.path.splitext(os.path.basename(filename))[0]
   
    #print('path {0}, file {1}'.format(path, filename))

    reader = pd.read_excel(source)
    output_name_template='{0}_{1}.xlsx'
    headers = None

    #file total file required to split
    total_records = len(reader)
    total_file_required = total_records/row_limit
    if total_file_required > int(total_file_required):
      total_file_required = int(total_file_required) + 1
    else:
      total_file_required = int(total_file_required)

    #assign header name use for create file
    if keep_headers:
        headers = reader.columns

    #generate file
    file_number = 1
    begin = 0
    end = row_limit
    while file_number <= total_file_required:
      #perform split file
      new_filename = None
      new_filename = output_name_template.format(base_filename,file_number)
      #print(new_filename)
      writer = pd.ExcelWriter(os.path.join(path,new_filename), engine='openpyxl')
      if total_file_required == file_number:
        excel = reader.iloc[begin:total_records]
        excel.to_excel(writer, 'raw_data', index=False)
      else:
        excel = reader.iloc[begin:end]
        excel.to_excel(writer, 'raw_data', index=False)

      begin = begin + row_limit
      end = end + row_limit

      writer.save()
      files_to_process.append(new_filename)
      file_number = file_number + 1
    return files_to_process

        
    
