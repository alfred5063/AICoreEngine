#!/usr/bin/python
# FINAL SCRIPT updated as of 3rd Dec 2020
# Workflow - STP-DM

# Declare Python libraries needed for this script
import pysftp
from utils.Session import session
from utils.guid import get_guid
from utils.audit_trail import audit_log
from utils.notification import send
from utils.logging import logging

myHostname = '115.164.158.48'
myUsername = "aan_prod"
myPassword = "TPA@3105"
myPort = "22"

def access_sftp(myHostname, myUsername, myPassword, myPort, base):

  error = []

  try:
    srv = pysftp.Connection(host = myHostname, username = myUsername, password = myPassword, private_key = 'C:/Users/alfred.simbun/.ssh/id_rsa')
    conn_flag = True
    print("Connection succesfully stablished ... ")
  except Exception as error:
    conn_flag = False
    print("Connection is failed ... ")
    pass

  if conn_flag == True:
    try:
      sftp.cwd('HOME/Others/NEWSYSTEM/20201118')
      remote_flag = True
    except:
      remote_flag = False
      print("Can't switch to remote directory ... ")
      pass

    if remote_flag == True:
      remoteFilePath = 'TUTORIAL.txt'
      localFilePath = r'C:\Users\sdkca\Desktop\TUTORIAL.txt'
      sftp.get(remoteFilePath, localFilePath)
    elif remote_flag == False:

  elif conn_flag == False:
    logging("STP-DM - Connection to SFTP failed.", error, stpdm_base)


    srv.close()







