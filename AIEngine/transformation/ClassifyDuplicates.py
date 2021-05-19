import pandas as pd
import math
from utils.audit_trail import audit_log
from utils.logging import logging

def classify_duplicates_from_merged(joined, source):
  """ Purpose: Get excel file source, return dataframe with non-duplicated and duplicated records
      Author: Promboon Jirawattanakitja
      Created on: 16/May/2019 """
  audit_log('classify_duplicates_from_merged', 'join excel value to find non-duplicated, duplicated or not found result.')
  try:
    non_duplicated = joined[joined['_merge'] == 'both'].drop(columns='_merge')
    duplicated = joined[joined['_merge'] == 'left_only'].drop(columns='_merge')
    not_found = joined[joined['_merge'] == 'right_only'].drop(columns='_merge')

    cols_to_drop = non_duplicated.columns.difference(source.columns)
    duplicated = duplicated.drop(columns=cols_to_drop)
    not_found = not_found.drop(columns=cols_to_drop)
  except Exception as error:
    logging('classify_duplicates_from_merged', error)

  return non_duplicated, duplicated, not_found

def convert_merge_to_code(row, column_to_validate):
  audit_log('convert_merge_to_code', 'merge action codes')
  try:
    conversion_dict = {'X': 2}
    value_place_dict = {
      'NRIC': int(math.pow(10, 0)),
      'OtherIC': int(math.pow(10, 1)),
      'Policy Num': int(math.pow(10, 2)),
      'Principal': int(math.pow(10, 3)),
      'Plan Expiry Date': int(math.pow(10, 4)),
      'DOB': int(math.pow(10, 5)),
      'Employee ID': int(math.pow(10, 6)),
      'Relationship': int(math.pow(10, 7)),
    }

    if row['_merge']  == 'both' or row['_merge'] == True:
      import_type = row['Import Type']
      value = conversion_dict[import_type] if import_type in conversion_dict else 1
      place_value = value_place_dict[column_to_validate]
      return value * place_value
  except Exception as error:
    logging('convert_merge_to_code', error)
  return 0

def get_matches_result(joined, column_to_validate, base):
  try:
    audit_log('get_matches_result', 'Convert matches to action codes', base)
    joined['_merge'] = joined.apply(lambda row: convert_merge_to_code(row, column_to_validate), axis=1)
    return joined
  except Exception as error:
    logging('get_matches_result', error, base)
    raise error

def remove_excess_column(raw_df, joined, base):
  try:
    audit_log('remove_excess_column', 'Remove all the temp columns', base)
    merge = joined['_merge']
    cols_to_drop = joined.columns.difference(raw_df.columns)
    clean_df = joined.drop(columns=cols_to_drop)
    clean_df['_merge'] = merge
    return clean_df
  except Exception as error:
    logging('remove_excess_column', error, base)
    raise error

#example code:
#source ='C:\\DM\\New\\0114-RHB Format - 1618.xlsx'
#destination = 'C:\\DM\\New\\converted.xlsx'
#non_duplicated, duplicated = classify_duplicates(source)
#export_excel(destination, non_duplicated, duplicated)
