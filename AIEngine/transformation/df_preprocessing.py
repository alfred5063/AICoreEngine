import pandas as pd
from utils.Session import session
from utils.audit_trail import audit_log
from utils.logging import logging
from pandas import Timestamp
from datetime import datetime

#def mark_as_unmatched(null_df, function_name):
#  null_df['matched_' + function_name] = 0
#  return null_df

def remove_null(df, column_for_checking, base):
  audit_log('remove_null', 'Remove null')
  try:
    invalid = df[column_for_checking].isna()
    valid = ~ invalid
    non_empty_df = df[valid] # ~ is used as 'not' for bitwise operation used in Dataframe
  except Exception as error:
    logging('remove_null', error, base)
  return non_empty_df, df[invalid]

def fix_dates(df, base):
  audit_log('fix_date', 'convert date format', base)
  for column in df.columns:
    try:
      df[column] = df[column].dt.strftime('%d/%m/%Y').values
    except Exception as error:
      pass
      #logging('fix_date', error)
  return df


def preprocessing(df, base, column_for_checking=None):
  try:
    audit_log('preprocessing', 'replace Na records with empty string', base)
    #return non_empty_df, null_df
    df.replace('-', None)
    df = df.where((pd.notnull(df)), None)
    return df
  except Exception as error:
    logging('preprocessing', error, base)
    raise error
