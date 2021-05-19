import json
import numpy as np
import pandas as pd
import xlrd
import os
import pymssql
from datetime import date
from connector.dbconfig import read_db_config
from pathlib import Path
from connector.connector import MSSqlConnector

#function to check file extension. Only support xls. and xlsx.and return file path
def excel_file_filter(filename, extensions=['.xls', '.xlsx']):
    return any(filename.endswith(e) for e in extensions)

#function to get file name
def get_filenames(root):
    filename_list = []
    for path, subdirs, files in os.walk(root):
        for filename in filter(excel_file_filter, files):
            filename_list.append(os.path.join(path, filename))
    return filename_list

#function to get today's date
def GETDATE():
    today = str(date.today())
    return today

#function to insert data into mssql by query
def insert_mssql(datajson, cmurl, year, cmemail):

    current_path = os.path.dirname(os.path.abspath(__file__))
    dti_path = str(Path(current_path).parent)
    config = read_db_config(dti_path+r'\config.ini', 'mssql')
    #config = read_db_config(r'C:\Users\NURFAZREEN.RAMLY\Source\Repos\rpa_engine\RPA.DTI\config.ini', 'mssql')
    #serv = config['server']
    #db = config['database']
    #userdb = config['user']
    #pwd = config['password']
    conn = MSSqlConnector()

    cursor = conn.cursor()
    cursor.execute(''' INSERT INTO [cba].[ops_disbursement_req_master]
                                                            ([content]
                                                            ,[year]
                                                            ,[createdby]
                                                            ,[modifiedby]
                                                            ,[createdon]
                                                            ,[modifiedon]
                                                            ,[url])
                                                        VALUES
                                                            (%s
                                                            ,%s
                                                            ,%s
                                                            ,%s
                                                            ,%s
                                                            ,%s
                                                            ,%s)
                    ''', (datajson, year, cmemail, cmemail, GETDATE(), GETDATE(), cmurl))
    conn.commit()
    conn.close()
