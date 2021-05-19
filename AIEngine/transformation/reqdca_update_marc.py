#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - REQ/DCA

# Declare Python libraries needed for this script
import pandas as pd
import numpy as np
import time
import psycopg2
import json
import copy
from connector.connector import MySqlConnector, MSSqlConnector
from connector.dbconfig import *
from automagic.automagica_reqdca import *
from datetime import datetime as dt
from utils.audit_trail import audit_log
from utils.logging import logging
import pyodbc as db
from loading.query.query_as_df import *

# Function to update document
def update_in_marc(newmain_df, doc_template, email, reqdca_base, document):

  try:

    # Get the current date and time
    named_tuple = time.localtime()
    current_date = time.strftime("%d-%m-%Y", named_tuple)
    current_time = time.strftime("%H%M", named_tuple)
    current_year = dt.now().year

    # Database connection
    conn = MSSqlConnector()
    cur = conn.cursor()

    # Iterate the case ID and create and save the DCA document for each ID
    i = 0
    for i in range(len(newmain_df)):
      try:

        # Prepare variables
        caseid = int(newmain_df.iloc[i]['Case ID'])
        attention = newmain_df.iloc[i]['Attention To']
        insurer = 'Insurer'
        address = newmain_df.iloc[i]['Address']

        if document == 'DCA':
          subcaseid = int(newmain_df.iloc[i]['Sub Case ID_x'])
          cln = str(newmain_df.iloc[i]['Disbursement Listing'])
          mytype = newmain_df.iloc[i]['Type_x']
          amount = newmain_df.iloc[i]['Total']
          HNS = newmain_df['Disbursement Claims'][i]
          remarks = newmain_df['Reason'][i]
        else:
          subcaseid = int(newmain_df.iloc[i]['Sub Case ID'])
          cln = str(newmain_df.iloc[i]['Client Listing Number'])
          mytype = newmain_df.iloc[i]['Type']
          amount = newmain_df.iloc[i]['Amount']
          HNS = newmain_df['HNS Number'][i]
          remarks = newmain_df['Reason'][i]

        # Updating record in MARC database
        if mytype == 'Admission':
          browser = login_Marc(reqdca_base)
          time.sleep(10)
          navigate_page(browser, 'inpatient', 'cases', 'inpatient_cases_search')
          time.sleep(10)
          myflag = search_Marc_inpatient(browser, caseid, reqdca_base)
          if myflag == '1':
            create_doc_inpatient(browser, attention, amount, remarks, HNS, doc_template, insurer, caseid, address, cln, current_date, current_time, reqdca_base)
          else:
            audit_log("REQDCA - DCA - MARC document is not updated. Case ID [ %s ] is locked." %caseid, "Completed...", reqdca_base)
            pass;
          time.sleep(5)

        else:
          browser = login_Marc(reqdca_base)
          time.sleep(10)
          navigate_page(browser, 'inpatient', 'csucases', 'inpatient_cases_csu_search')
          time.sleep(10)
          myflag = search_Marc_outpatient(browser, subcaseid, reqdca_base)
          if myflag == '1':
            create_doc_outpatient(browser, attention, amount, remarks, HNS, doc_template, insurer, subcaseid, address, cln, current_date, current_time, reqdca_base)
          else:
            audit_log("REQDCA - DCA - MARC document is not updated. Case ID [ %s ] is locked." %subcaseid, "Completed...", reqdca_base)
            pass;
          time.sleep(5)
          
        # Force pause to allow current process is completed (need to change this to wait based on process' time and not hardcoded)
        time.sleep(40)
        browser.close()
        
      except Exception as error:
        continue
    cur.close()
  except Exception as error:
    logging("REQ/DCA - extract_info", error, reqdca_base)
    pass
