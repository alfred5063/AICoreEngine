#Created by Yeoh Eik Den
#Created on 27th May 2019
#Descripton:
#Manage column is to select column, and remove column
#This transformation include replace value, rename column, change data type, pivot, unpivot, convert to list, group by, change order (asc, desc)
import pandas as pd
import os

#select all columns
def get_columns_name(df):
  data_frame = pd.DataFrame(df)
  columns_name = list(data_frame.columns)
  return columns_name

#rename column (single column)
def set_columns_rename(df, old_column_name, new_column_name):
  try:
    df.rename(columns = {old_column_name:new_column_name}, inplace=True)
  except Exception as e:
    return str(e.args)

  return df

#pivot data frame
def pivot(df, columns_name, values):
   data_frame = pd.pivot(df, columns=columns_name, values=values)
   return data_frame

#sorting single column
def sort(df, column_name, asc=True):
  df.sort_values(column_name,ascending=asc, inplace=True)
  return df

#remove columns mane = ['col1','col2', 'col3']
def remove_column(df, columns_name):
  df.drop(columns=columns_name, inplace=True)
  return df

#group by column name (single column)
def groupby(df, column_name):
  df.groupby(by=column_name)
  return df

#convert data frame to list
def df_to_list(df):
  list = df.values.tolist()
  return list

#replace value
#def replace_value(df, column_name, old_value, new_value):
#  df[column_name] = df[column_name].apply({old_value:new_value}.get)
#  return df

#def to_numeric(df, column_name):
#  df[column_name] = pd.to_numeric(df[column_name])
#  return df

#test script
#directory = "C:\\Asia-Assistance\\RPA3.0\\DM"
#for filename in os.listdir(directory):
#    if filename.endswith(".xlsx"):
#        df = pd.read_excel(os.path.join(directory, filename))
#        #df = pivot(df, "Import Type", "Member Full Name")
#        #df = remove_column(df, ["Member Full Name"])
#        #c_name = get_columns_name(df)
#        #df = sort(df, "Member Full Name")
#        #df = set_columns_rename(df, "Member Full Name","Principal Name")
#        #df = groupby(df, "Import Type")
#        replace_value(df, 'Import Type', 'X', 1)
#        replace_value(df, 'Import Type', 'N', 2)

#        #df = to_numeric(df, "No")
#        print(df)
#        continue
#    else:
#        continue
