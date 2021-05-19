#!/usr/bin/python
# FINAL SCRIPT updated as of 16th June 2020
# Version 1

# Declare Python libraries needed for this script
import pandas as pd
import os

#dirName = source
# Get list of filenames
def getListOfFiles(dirName):
    listOfFile = os.listdir(dirName)
    allFiles = list()
    for entry in listOfFile:
        fullPath = os.path.join(dirName, entry)
        if os.path.isdir(fullPath):
            #allFiles = allFiles + getListOfFiles(fullPath)
            pass
        else:
            allFiles.append(fullPath)
    return allFiles


# Get filename based on client name
def search_cmlfile(jsondf, cmlfiles, current_year):
  
  temp_df2 = pd.DataFrame()
  finalpath_df = pd.DataFrame()
  finalpath_df2 = pd.DataFrame()
  clientlist = pd.DataFrame()
  filelist = list(getListOfFiles(cmlfiles))

  i = 0
  for i in range(len(jsondf['Client'])):
    client = jsondf['Client'][i].split(" ")
    mylist = list(filter(lambda x: '\\%s\\' %str(client[i]) in x, filelist))
    #mylist = list(filter(lambda x: '\\%s\\' %str(jsondf['Client'][i]) in x, filelist))
    mylist = list(filter(lambda x: '%s' %str(current_year) in x, mylist))
    mylist = pd.DataFrame(list(filter(lambda x: '.xlsx' in x, mylist)), columns = ['Filenames'])

    a = 0
    for a in range(len(client)):
      temp_df = pd.DataFrame(columns = ['Filenames'])
      searched = mylist[mylist['Filenames'].str.lower().astype(str).str.contains(client[a].lower())]
      if searched.empty != True:
        temp_df = temp_df.append(searched)
        temp_df = temp_df.drop_duplicates().reset_index().drop(columns = ['index'])
        for b in range(len(temp_df['Filenames'])):
          temp_df.loc[b, 'Client Name'] = str(jsondf['Client'][i])
        mylist = searched
        temp_df2 = temp_df2.append(temp_df)
        finalpath_df = temp_df2.drop_duplicates()
      else:
        pass
      continue
    finalpath_df2 = finalpath_df2.append(finalpath_df)

  finalpath_df2 = finalpath_df2.drop_duplicates().reset_index().drop(columns = ['index'])
  return finalpath_df2


