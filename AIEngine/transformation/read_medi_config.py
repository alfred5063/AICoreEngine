from configparser import ConfigParser
import os

#""" Read database configuration file and return a dictionary object
#:param filename: name of the configuration file
#:param section: section of database configuration
#:return: a dictionary of database parameters
#"""
def read_config(filename, section):
    # create parser and read ini configuration file
    parser = ConfigParser()
    parser.read(filename)
    #print('parser')
    #print(parser)

    db = {}
    if parser.has_section(section):
        #print('FOUND SECTION')
        items = parser.items(section)
        for item in items:
            db[item[0]] = item[1]
    else:
        raise Exception('{0} not found in the {1} file'.format(section, filename))

    return db

#example - test function
#msg = read_db_config('C:\Asia-Assistance\RPA3.0\RPA\Master\RPA.DTI\extraction\marc\config.ini', 'mysql')
#print(msg)
current_path = os.getcwd()


def get_skiprows(bdx_type):
  value = None
  config = read_config(current_path+r'\config.ini', 'medi_bdx')
  value = int(config[bdx_type])
  return value

def get_env():
  value = None
  config = read_config(current_path+r'\config.ini', 'medi_file')
  value = config["env"]
  return value

