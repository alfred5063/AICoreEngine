from configparser import ConfigParser
import os
from utils.audit_trail import audit_log
from utils.logging import logging
#""" Read database configuration file and return a dictionary object
#:param filename: name of the configuration file
#:param section: section of database configuration
#:return: a dictionary of database parameters
#"""
def read_db_config(filename, section, base):
    audit_log('read_db_config', 'read db configuration setting.', base)
    # create parser and read ini configuration file
    try:
      parser = ConfigParser()
      parser.read(filename)
      #print('parser')
      #print(parser)

      # get section, default to mysql
      db = {}
      if parser.has_section(section):
          #print('FOUND SECTION')
          items = parser.items(section)
          for item in items:
              db[item[0]] = item[1]
      else:
          raise Exception('{0} not found in the {1} file'.format(section, filename))
    except Exception as error:
      logging('read_db_config', error, base)

    return db

#example - test function
#msg = read_db_config('C:\Asia-Assistance\RPA3.0\RPA\Master\RPA.DTI\extraction\marc\config.ini', 'mysql')
#print(msg)
