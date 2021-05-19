#!/usr/bin/python
# FINAL SCRIPT updated as of 15th April 2021
# Workflow - Finance SOA
# Version 1

# Declare Python libraries needed for this script
import pandas as pd
import numpy as np
import datetime as dt

#check column use to compare date
def checkbilldate(soa_df,len_payment_date):
  datetocompare = ['Document Date', 'Invoice Date', 'Date']
  for i in range(len(datetocompare)):
    try:
      billDate = soa_df.iloc[len_payment_date][datetocompare[i]]
      kwrd = datetocompare[i]
    except Exception as error:
      pass
    continue
  return billDate

#count overdue date
def date_due(Date, day):
  date_pattern = get_date(Date)
  date_1 = dt.datetime.strptime(str(Date), date_pattern)
  end_date = date_1 + dt.timedelta(days=day)
  overdue_date = end_date.strftime("%d/%m/%Y")

  return overdue_date, end_date
  
#check date pattern
def get_date(billDate):
    date_patterns = ["%d-%m-%Y", "%Y-%m-%d", "%d-%m-%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y %H:%M:%S.f", "%Y-%m-%d %H:%M:%S.%f", "%Y/%m/%d", "%d/%m/%Y", "%Y/%m/%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S.%f", "%Y/%m/%d %H:%M:%S.%f"]
    for pattern in range(len(date_patterns)):
      try:
          dt.datetime.strptime(str(billDate), date_patterns[pattern])
          date_pattern = date_patterns[pattern]
      except:
        pass
      continue
    return date_pattern

def check_due_date(soa_df, vt_df):

  # List for filter(Client)test
  due30_days = ['Adventist Hospital & Clinic Services (M)', 'Hospital Lam Wah Ee', 'Island Hospital Sdn Bhd', 'Prince Court Medical Centre Sdn Bhd']
  due45_days = ['Ara Damansara Medical Centre Sdn Bhd', 'Hospital Fatimah', 'Mount Miriam Cancer Hospital', 'RSD Hospitals Sdn Bhd - PMC', 'RSD Hospitals Sdn Bhd - SJMC']

  # get to know is which hospital (need to know the correct place)
  hospital = vt_df.iloc[0]['Doc Types']

  # drop the client is empty
  #soa_df['Client'].replace('', np.nan, inplace=True)
  #soa_df.dropna(subset=['Client'], inplace=True)
  #soa_df = soa_df.reset_index(drop=True)

  # calculate overdue date
  len_payment_date = 0
  for len_payment_date in range(soa_df.shape[0]):
    #print('soa_df[Payment Date].iloc[len_payment_date]')
    #print(pd.isnull(soa_df.iloc[len_payment_date]['Payment Date']) == True)
    #print('soa_df[Client].iloc[len_payment_date] != ""')
    #print(soa_df.iloc[len_payment_date]['Client'] != "")
    if pd.isnull(soa_df.iloc[len_payment_date]['Payment Date']) == True and pd.isnull(soa_df.iloc[len_payment_date]['Client']) == False:
     
      if hospital in due30_days:
        # 30 days function
        billDate = checkbilldate(soa_df,len_payment_date)   
        overdue_date, end_date = date_due(billDate,30)


        # put remark based on overdue date
        current_date = dt.datetime.now()

        # key in overdue date
        overdue_day = current_date - end_date
        soa_df.loc[len_payment_date,'Overdue Day'] = overdue_day
        # get CML amount
        amount = soa_df.iloc[len_payment_date]['Insurance']
        if amount == "0" or amount == "" or amount == 0 or pd.isnull(amount):
          if current_date > end_date:
            soa_df.loc[len_payment_date,'Remarks'] = "pending for reimbursement from client"
          else:
            soa_df.loc[len_payment_date,'Remarks'] = "not due yet"
        else:
          soa_df.loc[len_payment_date,'Remarks'] = "payment received from client"

      elif hospital in due45_days:
        # 45 days function
        
        billDate = checkbilldate(soa_df,len_payment_date)      
        overdue_date, end_date = date_due(billDate,45)


        # put remark based on overdue date
        current_date = dt.datetime.now()

        # key in overdue date
        overdue_day = current_date - end_date
        soa_df.loc[len_payment_date,'Overdue Day'] = overdue_day
        # get CML amount
        amount = soa_df.iloc[len_payment_date]['Insurance']
        if amount == "0" or amount == "" or amount == 0 or pd.isnull(amount):
          if current_date > end_date:
            soa_df.loc[len_payment_date,'Remarks'] = "pending for reimbursement from client"
          else:
            soa_df.loc[len_payment_date,'Remarks'] = "not due yet"
        else:
          soa_df.loc[len_payment_date,'Remarks'] = "payment received from client"

      else:
        # 60 days function
        billDate = checkbilldate(soa_df,len_payment_date)
        overdue_date, end_date = date_due(billDate,60)


        # put remark based on overdue date
        current_date = dt.datetime.now()

        # key in overdue date
        overdue_day = current_date - end_date
        print(overdue_day.days)
        soa_df.loc[len_payment_date,'Overdue Day'] = overdue_day.days
        # get CML amount
        amount = soa_df.iloc[len_payment_date]['Insurance']

        if amount == "0" or amount == "" or amount == 0 or pd.isnull(amount):
          if current_date > end_date:
            soa_df.loc[len_payment_date,'Remarks'] = "pending for reimbursement from client"
          else:
            soa_df.loc[len_payment_date,'Remarks'] = "not due yet"
        else:
          soa_df.loc[len_payment_date,'Remarks'] = "payment received from client"
    else:
      pass
  soa_df['Overdue Day'] = soa_df['Overdue Day'].replace(np.nan, 0.0, regex=True)
  soa_df['Overdue Day'] = soa_df['Overdue Day'].astype(int)
  return soa_df





