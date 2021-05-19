import pandas as pd

def clear_matrix_code_with_condition(cleaned_df):
  dm_df = cleaned_df
  dm_df.columns = dm_df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
  #print(list(dm_df.columns))
  dm_df.loc[dm_df.policy_num.isin(['nan']), '_merge'] = '0'
  dm_df.loc[dm_df.relationship.isin(['nan']), '_merge'] = '0'
  dm_df.loc[dm_df.dob.isin(['nan']), '_merge'] = '0'
  dm_df.loc[dm_df.plan_expiry_date.isin(['nan']),'_merge']='0'
  return dm_df
