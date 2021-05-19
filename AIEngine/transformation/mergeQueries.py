#Merge Queries include 2 different data frame or table to merge either left, right, union, or inner
import pandas as pd
from utils.audit_trail import audit_log

def merge(df1, df2, base, **kwargs):
  audit_log('Processing data' , 'combine MARC and Excel data', base)
  joined = pd.merge(df1, df2, **kwargs)
  return joined

def left_join(df1, df2, column_name):
  joined = pd.merge(df1, df2, how='left', on=column_name)
  return joined

def right_join(df1, df2, column_name):
  joined = pd.merge(df1,df2, how='right', on=column_name)
  return joined

def inner_join(df1, df2, column_name):
  joined = pd.merge(df1, df2, how='inner', on=column_name)
  return joined

def outer_join(df1, df2, column_name):
  joined = pd.merge(df1,df2, how='outer', on=column_name)
  return joined
