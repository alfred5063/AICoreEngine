#!/usr/bin/python
# FINAL SCRIPT updated as of 6th September 2019
# Scriptor - Leow Jun Shou

# Declare Python libraries needed for this script
import urllib.request
import json
from pandas.io.json import json_normalize
import pandas as pd

# Function to convert JSON URL/String to python dataframe
def json_To_Df(url = None, string = None):

  if string is not None:
    data = json.loads(string)

  if url is not None:
    response = urllib.request.urlopen(url)
    data = json.loads(response.read())

  dataf = json_normalize(data).transpose()

  return dataf

# For MSIG, get the execution_parameters only
def get_execution_parameters(dataframe, number):

  dataf = dataframe
  execution_parameters = (dataf.loc["execution_parameters"])[0]
  execution_parameters = execution_parameters[number]
  adm_id = execution_parameters["Adm. No"]
  case_id = execution_parameters["CASE ID"]
  ch = execution_parameters["CH"]
  return adm_id, case_id, ch


#json_To_Df(url='https://jsonplaceholder.typicode.com/todos/1')

#a={
#    "guid":"d",
#    "taskid":"d",
#    "stepname": "d",
#    "email": "d",
#    "execution_parameters":[{"case_id":1,
#                          "disbursement_id":21323,
#                          "type": "texbox"},
#                         {"texbox":2,
#                          "disbursement_id":2131,
#                          "type":"textbox"}
#                        ],
#    "execution_parameters_2":[{ 
#                            }],
#    "authentication":{ "username": "testae", 
#                     "password": "wrerwerwerwerwe",
#                     "privatekey": "werewerweorwe"
#                    }
#}


