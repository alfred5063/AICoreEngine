import pandas as pd
from pandas import ExcelWriter
import os

#  """ Purpose: This function is use for directory create json file
#      Author: Yeoh Eik Den
#      Created on: 17 May 2019"""

def createJson(filename, value):
  df = pd.DataFrame(value)
  export = df.to_json(filename)
  print("Json file created")

#example
#filepath = "C:\\Asia-Assistance\\DM\\new\\"
#filename="test.json"
#value={'a':[1,3,5,7,4,5,6,4,7,8,9],'b':[3,5,6,2,4,6,7,8,7,8,9]}
#createJson(filepath+filename, value)

