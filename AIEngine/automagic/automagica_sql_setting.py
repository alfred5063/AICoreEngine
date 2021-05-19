from connector.connector import MSSqlConnector
from utils.audit_trail import audit_log
from utils.logging import logging
from loading.query.query_as_df import mssql_get_df_by_query_without_base
from utils.application import *

def get_application_url(application_name, url=None):
   connector = MSSqlConnector
   query = '''SELECT [application_id]
                    ,[application_url]
               FROM [dbo].[application]
	             WHERE [application_name] = %s'''
   params = (application_name)
   result_df = mssql_get_df_by_query_without_base(query, params, connector)

   app_info = browser()
   for index, app_item in result_df.iterrows():
    app_info.application.appId = app_item["application_id"]
    app_info.application.appName = application_name
    if url!=None:
      app_info.application.appUrl = app_item["application_url"].replace('https://marcmy.asia-assistance.com/marcmy/', url)
    else:
      app_info.application.appUrl = app_item["application_url"]

   #print(app_info.application.appUrl)
   return app_info

def get_rpa_application_url(application_name):
   connector = MSSqlConnector
   query = '''SELECT [application_id]
                    ,[application_url]
               FROM [dbo].[application]
	             WHERE [application_name] = %s'''
   params = (application_name)
   result_df = mssql_get_df_by_query_without_base(query, params, connector)
   id = ""
   url =""

   for index, app_item in result_df.iterrows():
     url = app_item["application_url"]
     id = app_item["application_id"]

   return id, url

def get_xpath_by_key(applicationId, key):
   connector = MSSqlConnector
   query = '''SELECT [xpath]
                    ,[x_coordinate]
                    ,[y_coordinate]
               FROM [dbo].[application_details]
	             WHERE [application_id] = %s and [key] = %s'''
   params = (applicationId, key)
   result_df = mssql_get_df_by_query_without_base(query, params, connector)

   xpath = None
   for index, app_item in result_df.iterrows():
    xpath = app_item["xpath"]
  
   return xpath

def get_xpath_list(applicationId):
   connector = MSSqlConnector
   query = '''SELECT [key],[xpath]
               FROM [dbo].[application_details]
	             WHERE [application_id] = %s'''
   params = (applicationId)
   result_df = mssql_get_df_by_query_without_base(query, params, connector)  
   return result_df

def get_xpath_coordinate_by_key(applicationId, key):
   connector = MSSqlConnector
   query = '''SELECT [xpath]
                    ,[x_coordinate]
                    ,[y_coordinate]
               FROM [dbo].[application_details]
	             WHERE [application_id] = %s and [key] = %s'''
   params = (applicationId, key)
   result_df = mssql_get_df_by_query_without_base(query, params, connector)

   x = None
   y = None
   for index, app_item in result_df.iterrows():
    x = app_item["x_coordinate"]
    y = app_item["y_coordinate"]
  
   return x, y

