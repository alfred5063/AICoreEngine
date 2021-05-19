#
#
#
#

import psycopg2
import pymssql
import mysql.connector
from mysql.connector import Error
from connector.dbconfig import read_db_config
import os
from pathlib import Path

current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent)

def PostgresConnector(config=None):
  """ Connect to PostgreSQL database (test) """
  #os.chdir("")

  try:
    conn = None
    if not config:
      config = read_db_config(dti_path+r'\config.ini', 'postgresql')

    conn = psycopg2.connect(user = config['uid'],
                                  password = config['password'],
                                  host = config['server'],
                                  port = config['port'],
                                  database = config['database'])
    return conn
  except (Exception, psycopg2.Error) as error :
    print ("Error while connecting to PostgreSQL", error)

def MySqlConnector(config=None, host=False):
  """ Connect to MySQL database """
 
  try:
    conn = None
    if not config:
      config = read_db_config(dti_path+r'\config.ini', 'mysql')
    print('Connecting to MARC database...')
    conn = mysql.connector.connect(host=config['host'],database=config['database'],
                                       user=config['user'], password=config['password'])

    if conn.is_connected():
      #print('connection established.')
      return conn
    else:
        print('connection failed.')
  except Error as error:
    print(error)

def MySqlConnector_Prod2(config=None, host=False):
  """ Connect to MySQL database """
 
  try:
    conn = None
    if not config:
      config = read_db_config(dti_path+r'\config.ini', 'mysql_prod2')
    print('Connecting to MARC database...')
    conn = mysql.connector.connect(host=config['host'],database=config['database'],
                                       user=config['user'], password=config['password'])

    if conn.is_connected():
      #print('connection established.')
      return conn
    else:
        print('connection failed.')
  except Error as error:
    print(error)

def MSSqlConnector(config=None, host=False):
  """ Connect to MSSQL database """

  try:
    conn = None
    if not config:
      config = read_db_config(dti_path+r'\config.ini', 'mssql')
      #conn = pymssql.connect(server=config['server'], user=config['user'], password=config['password'], database=config['database'])
      #print(conn)
      #window authentication connection
      conn = pymssql.connect(server=config['server'], database=config['database'])
    if conn.cursor():
      #print('connection established.')
      return conn
    else:
        print('connection failed.')
  except Error as error:
    print(error)

