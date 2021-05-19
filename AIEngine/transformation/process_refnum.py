#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - MSIG

# Declare Python libraries needed for this script
import sys
import json
import pandas as pd
import numpy as np
import os
import shutil
from datetime import datetime
from connector.connector import MySqlConnector
from utils.logging import logging

# Function to process input files from the user
def process_refnum(jsondf, msig_base, preview):

  try:
    
    # Create temporary empty dataframes
    validated_df = pd.DataFrame(columns = ['Adm. No', 'CASE ID', 'CH', 'Status'])
    nonvalidated_df = pd.DataFrame(columns = ['Adm. No', 'CASE ID', 'CH', 'Status'])
    maindf = pd.DataFrame(columns = ['Adm. No', 'CASE ID', 'CH'])

    # Prepare main dataframe.
    # Segregate rows with missing elements from the ones with elements.
    # Include Client Ref.Num and Status columns
    jsondf = jsondf.drop_duplicates()
    jsondf = jsondf.replace('N/A', ' ')
    maindf = jsondf[(jsondf != 0).all(1)]
    maindf.insert(3, 'Client Ref.Num', maindf['CASE ID'].map(str) + " " + maindf['CH'].map(str))

    # Prepare list of Case ID (Adm. No) for preprocessing tasks
    # Check if Adm. No in user's file matches the same in MARC
    adm_id = []
    adm_id_arrayed = np.array(maindf['Adm. No'])
    if preview != 'False':
      # create a cursor for MARC database connection
      conn = MySqlConnector()
      cur = conn.cursor()
      #print("test1")
      #i = 0
      for i in range(len(adm_id_arrayed)):
        adm_id = adm_id_arrayed[i]
        adm_id

        # Calling Stored Procedure
        paramater = [str(adm_id),]
        stored_proc = cur.callproc('msig_query_marc_for_insurer_details', paramater)
        for i in cur.stored_results():
            results = i.fetchall()

        fetched_result = pd.DataFrame(results)

        if fetched_result.empty != True:
          validated_df = validated_df.append(maindf.loc[maindf['Adm. No'].isin([str(adm_id)])], ignore_index = True, sort = False)
          validated_df['Status'] = "Record is valid based on Adm. ID. Can be processed."
        else:
          nonvalidated_df = nonvalidated_df.append(maindf.loc[maindf['Adm. No'].isin([str(adm_id)])], ignore_index = True, sort = False)
          nonvalidated_df['Status'] = "Case ID is NOT FOUND in MARC. Please check. Consider to process?"

      # Consolidated two dataframes for preview and keep in JSON format
      maindf = pd.DataFrame(validated_df.append(nonvalidated_df, ignore_index = True, sort = False))

      # close the communication with the MARC database
      cur.close()
      conn.close()

    return maindf

  except Exception as error:
    logging("MSIG - process_refnum", error, msig_base)
