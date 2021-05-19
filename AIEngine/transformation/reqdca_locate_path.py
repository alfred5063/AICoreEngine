#!/usr/bin/python
# FINAL SCRIPT updated as of 8th July 2020
# Workflow - REQ/DCA

# Declare Python libraries needed for this script
import glob
import json
import pandas as pd


def locate_path(dcamain_df):

  # Read the current JSON string containing file paths

  # For testing purpose
  #myurl = r'C:\Users\adm.alfred.simbun\OneDrive - AXA\ALFRED - Development\TESTING\REQDCA_Paths\Testing_Path_DataFrame.json'

  # For Production
  myurl = r'\\Dtisvr2\cba_uat\Result\REQDCA_Paths\Testing_Path_DataFrame.json'

  with open(myurl, 'r') as f:
    path_df = json.load(f)
    path_df = pd.DataFrame.from_dict(path_df, orient='index')

  # Extract out member's reference ID
  memExtID = dcamain_df['Member Ref. ID']

  # Check if file path exist in path_df dataframe based on member's reference ID
  i = 0
  temp_df = pd.DataFrame()
  for i in range(len(memExtID)):
    extracted_path = path_df.loc[path_df['Ext Ref'] == memExtID[i]]
    extracted_path = extracted_path[['Ext Ref', 'Path']].reset_index()
    del extracted_path['index']
    extracted_path = extracted_path.drop_duplicates()
    rownum = extracted_path['Ext Ref'].count()
    if rownum == 1:
      temp_df = temp_df.append(extracted_path, ignore_index = True, sort = False)
      temp_df = temp_df.drop_duplicates()
    else:
      extracted_path.loc[i, 'Ext Ref'] = "%s" % memExtID[i]
      extracted_path.loc[i, 'Path'] = str("Path not found")
      temp_df = temp_df.append(extracted_path, ignore_index = True, sort = False)
      temp_df = temp_df.drop_duplicates()

  # Merge two dataframes and parse to the workflow script
  df2 = pd.merge(left = dcamain_df, right = temp_df[['Ext Ref', 'Path']], left_on = dcamain_df['Member Ref. ID'], right_on = temp_df['Ext Ref'])
  df2 = df2.drop(columns = ['key_0', 'Ext Ref'])

  # Return the dataframe to the workflow script
  return df2


