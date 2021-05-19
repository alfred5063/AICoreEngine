import os
import time
from datetime import datetime
from directory.get_filename_from_path import get_file_name
from utils.audit_trail import audit_log
from utils.logging import logging
from shutil import copyfile
#""" Purpose: This class is use for directory create folder
#      Author: Yeoh Eik Den
#      Created on: 17 May 2019
#      Updated on: 29th May 2020"""

def movefile(source, destination, base):
  audit_log('movefile', 'Moving file to %s' % destination, base)
  try:
    os.rename(source, destination)
    print("Successfully moved file into %s" % destination)
  except FileNotFoundError as error:
    logging('movefile', error, base)
    print(error)

def move_file_to_archived(source, base):
  try:
    audit_log('move_file_to_archived', 'input is rename with timestamp and moved to Archived folder', base)
    path, file = get_file_name(source)
    now = datetime.now().strftime('%d%m%Y%H%M%S')
    timestamp = str(now)

    filename = file.split('.')[0]
    extension = file.split('.')[1]
    destination_filename = filename + '_' + timestamp + '.' + extension
    #to create archive forlder
    pathIndex = path.rfind("\\")
    pre_achive_path = path[:pathIndex]
    dirArchive = pre_achive_path+"\\Archived"

    if not os.path.exists(dirArchive):
      os.mkdir(dirArchive)

    destination = dirArchive+'\\{}'.format(destination_filename)
    movefile(source, destination, base)
    return destination
  except Exception as error:
    logging('move_file_to_archived', error, base)
    raise error

def move_to_archived(source):
  try:
    path, file = get_file_name(source)
    now = datetime.now().strftime('%d%m%Y%H%M%S')
    timestamp = str(now)

    filename = file.split('.')[0]
    extension = file.split('.')[1]
    destination_filename = filename + '_' + timestamp + '.' + extension
    #to create archive forlder
    pathIndex = path.rfind("\\")
    pre_achive_path = path[:pathIndex]
    dirArchive = pre_achive_path+"\\Archived"

    if not os.path.exists(dirArchive):
      os.mkdir(dirArchive)

    destination = dirArchive+'\\{}'.format(destination_filename)
    os.rename(source, destination)
    return destination
  except Exception as error:
    logging('move_file_to_archived', error, base)
    raise error

def copy_file_to_archived(source):
  try:
    path, file = get_file_name(source)
    now = datetime.now().strftime('%d%m%Y%H%M%S')
    timestamp = str(now)

    #filename = file.split('.')[0]
    #extension = file.split('.')[1]
    #destination_filename = filename + '_' + timestamp + '.' + extension
    #to create archive forlder
    pathIndex = path.rfind("\\")
    pre_achive_path = path[:pathIndex]
    dirArchive = pre_achive_path+"\\Archived\\"+now
    
    if not os.path.exists(dirArchive):
      os.mkdir(dirArchive)

    copyfile(source, dirArchive+'\\{}'.format(file))
    destination = dirArchive+'\\{}'.format(file)
    return destination
  except Exception as error:
    print('Error: {0}'.format(error))
    raise error

def move_file_to_result_medi(source):
  try:
    path, file = get_file_name(source)
    now = datetime.now().strftime('%d%m%Y%H%M%S')
    timestamp = str(now)

    #filename = file.split('.')[0]
    #extension = file.split('.')[1]
    #destination_filename = filename + '_' + timestamp + '.' + extension
    #to create archive forlder
    pathIndex = path.rfind("\\")
    pre_achive_path = path[:pathIndex]
    dirResult = pre_achive_path+"\\Result\\"+now
    
    if not os.path.exists(dirResult):
      os.mkdir(dirResult)

    destination = dirResult+'\\{}'.format(file)
    os.rename(source, destination)
    
    return destination
  except Exception as error:
    print('Error: {0}'.format(error))
    raise error


def move_file_to_result_dm(source, base):
  try:
    audit_log('move_file_to_archived', 'input is rename with timestamp and moved to Result folder', base)
    path, file = get_file_name(source)
    now = datetime.now().strftime('%d%m%Y%H%M%S')
    timestamp = str(now)

    filename = file.split('.')[0]
    extension = file.split('.')[1]

    destination_filename = filename + '_' + timestamp + '.' + extension
    #to create archive forlder
    pathIndex = path.rfind("\\")
    pre_result_path = path[:pathIndex]
    dirResult = pre_result_path + "\\Result"

    if not os.path.exists(dirResult):
      os.mkdir(dirResult)

    destination = dirResult+'\\{}'.format(destination_filename)
    movefile(source, destination, base)
    return destination
  except Exception as error:
    logging('move_file_to_result', error, base)
    raise error


def move_file_to_result(source, base):
  try:
    audit_log('move_file_to_archived', 'input is rename with timestamp and moved to Result folder', base)
    path, file = get_file_name(source)
    now = datetime.now().strftime('%d%m%Y%H%M%S')
    timestamp = str(now)

    filename = file.split('.')[0]
    extension = file.split('.')[1]

    destination_filename = filename + '_' + timestamp + '.' + extension
    #to create archive forlder
    pathIndex = path.rfind("\\")
    pre_result_path = path[:pathIndex]
    dirResult = pre_result_path + "\\Result"

    if not os.path.exists(dirResult):
      os.mkdir(dirResult)

    destination = dirResult+'\\{}'.format(destination_filename)
    movefile(source, destination, base)
    return destination
  except Exception as error:
    logging('move_file_to_result', error, base)
    raise error


def move_file_to_result_ar(source, base):
  try:
    audit_log('move_file_to_archived', 'input is rename with timestamp and moved to Result folder', base)
    path, file = get_file_name(source)
    now = datetime.now().strftime('%d%m%Y%H%M%S')
    timestamp = str(now)

    filename = file.split('.')[0]
    extension = file.split('.')[1]

    destination_filename = filename + '.' + extension
    #to create archive forlder
    pathIndex = path.rfind("\\")
    pre_result_path = path[:pathIndex]
    dirResult = pre_result_path + "\\Result"

    if not os.path.exists(dirResult):
      os.mkdir(dirResult)

    destination = dirResult+'\\{}'.format(destination_filename)
    movefile(source, destination, base)
    return destination
  except Exception as error:
    logging('move_file_to_result', error, base)
    raise error
