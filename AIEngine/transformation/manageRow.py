#Manage Row is to remove blank row, remove top, bottom or alternative row, remove duplicate.
import os
import pandas as pd
#remove bottom rows
def remove_bottom_row(df, bottom_number):
  df.drop(df.tail(bottom_number).index, inplace=True)
  return df

#remove top rows
def remove_top_row(df, top_number):
  df.drop(df.head(top_number).index, inplace=True)
  return df

#remove alternate rows
def remove_alternate_row(df, num=2):
  df = df.iloc[::num]
  return df

#remove duplicate
def remove_duplicate(df, column, **kwargs):
  df = df.drop_duplicates(column, **kwargs)
  return df


