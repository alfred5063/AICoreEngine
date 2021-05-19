# Check if folder is empty or not empty
from directory.files_listing_operation import getListOfFiles
from utils.logging import logging

def check_emptiness(input, base):
  input_flag = []
  new_files = []
  try:
    new_files = getListOfFiles(str(input))
    if new_files != []:
      input_flag = True
    else:
      input_flag = ''
  except Exception as error:
    logging("Error in detecting input files from source.", error, base)

  return input_flag, new_files
