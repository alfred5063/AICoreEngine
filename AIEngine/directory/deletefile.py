import os
from utils.logging import logging
from utils.audit_trail import audit_log
#""" Purpose: This function is use for delete file
#      Author: Yeoh Eik Den
#      Created on: 17 May 2019"""

def deletefile(filename, base):
  audit_log('deletefile', 'Delete file.')
  try:
    os.remove(filename)
    message = "Successfully delete file into %s" % filename
    print(message)
    return
  except FileNotFoundError as error:
    logging('deletefile', error, base)
    print(error)

#example code:
#filename = "C:\\Asia-Assistance\\DM\\new\\0114-RHB Format - 1618 - Copy.xlsx"
#deletefile(filename)
