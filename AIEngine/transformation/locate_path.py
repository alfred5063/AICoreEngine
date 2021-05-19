#!/usr/bin/python
# FINAL SCRIPT updated as of 4th October 2019
# Workflow - REQ/DCA
# Version 1

# Declare Python libraries needed for this script
import glob
from testing.shah.read_json import json_to_df
import pandas as pd

def locate_path(newmain_df):

  # Read the current JSON string containing file paths
  path_df = json_to_df(r'C:\Users\adm.alfred.simbun\Source\Repos\RoboticProcessAutomation\RPA.DTI\TESTING\input\Path_DataFrame.json')

  # Extract out member's reference ID
  memExtID = newmain_df['Member Ref. ID']

  # Check if file path exist in path_df dataframe based on member's reference ID
  i = []
  temp_df = pd.DataFrame()
  for i in range(len(memExtID)):
    extracted_path = path_df.loc[path_df['Ext Ref'] == '%s' % memExtID[i]]
    extracted_path = extracted_path[['Ext Ref', 'Path']].reset_index()
    del extracted_path['index']
    extracted_path = extracted_path.drop_duplicates()
    rownum = extracted_path['Ext Ref'].count()
    if rownum == 1:
      temp_df = temp_df.append(extracted_path, ignore_index = True, sort = False)
      temp_df = temp_df.drop_duplicates()
    elif rownum == 0:
      extracted_path['Path'] = "Path not found"
      temp_df = temp_df.append(extracted_path, ignore_index = True, sort = False)
      temp_df = temp_df.drop_duplicates()

  # Merge two dataframes and parse to the workflow script
  df2 = pd.merge(left = newmain_df, right = temp_df[['Ext Ref', 'Path']], left_on = newmain_df['Member Ref. ID'], right_on = temp_df['Ext Ref'])
  df2 = df2.drop(columns = ['key_0', 'Ext Ref'])

  # Return the dataframe to the workflow script
  return df2


