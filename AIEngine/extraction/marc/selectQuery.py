from mysql.connector import MySQLConnection, Error
from utils.audit_trail import audit_log
from utils.logging import logging
import mysql.connector
from mysql.connector import errorcode
def getmember(base):
    audit_log('getmember', 'Select member from MARC', base)
    try:
        conn = mysql.connector.connect(host=db_config['host'],database=db_config['database'],
                                       user=db_config['user'], password=db_config['password'])
        cursor = mySQLConnection.cursor(prepared=True)
        sql_select_query = """select * from member where mmid = 1 %s"""
        cursor.execute(sql_select_query, (MMID, ))
        record = cursor.fetchall()
        for row in record:
            print("mmid= ", row[0], )
    except mysql.connector.Error as error:
        print("Failed to get record from database: {}".format(error))
        logging('getmember', error, base)
    finally:
        # closing database connection.
        if (mySQLConnection.is_connected()):
            cursor.close()
            mySQLConnection.close()
            print("connection is closed")
