class BaseConfig:
  # General
  FLASK_APP = 'flaskr'
  # DB
  SQLALCHEMY_TRACK_MODIFICATIONS = True

class DevelopmentConfig(BaseConfig):
  ENV = 'development'
  FLASK_ENV = 'development'
  DEBUG = True
  TESTING = True
  #SQLALCHEMY_DATABASE_URI = 'mssql+pymssql://10.147.78.70/asia_assistance_rpa'
  SQLALCHEMY_DATABASE_URI = 'mssql+pymssql://10.147.49.43/asia_assistance_rpa'
  SECRET_KEY = 'RPA_SECRET' # os.urandom(24)
  #SERVER_NAME = 'localhost:5555'
  SERVER_NAME = '10.147.49.43:5555'
  
class LocalConfig(DevelopmentConfig):
  ENV = 'local'
  FLASK_ENV = 'local'
  DEBUG = True
  TESTING = True
  SQLALCHEMY_DATABASE_URI = 'mssql+pymssql://10.147.49.43/asia_assistance_rpa'
  SERVER_NAME = 'localhost:5555'

class UATConfig(BaseConfig):
  ENV = 'UAT'
  FLASK_ENV = 'UAT'
  DEBUG = False
  TESTING = False
  #SQLALCHEMY_DATABASE_URI = 'mssql+pymssql://10.147.78.70/asia_assistance_rpa_uat'
  SQLALCHEMY_DATABASE_URI = 'mssql+pymssql://10.147.49.43/asia_assistance_rpa_uat'
  SECRET_KEY = 'RPA_UAT_SECRET'
  #SERVER_NAME = '10.147.78.70:5556'
  SERVER_NAME = '10.147.49.43:5560'

class ProductionConfig(BaseConfig):
  ENV = 'production'
  FLASK_ENV = 'production'
  DEBUG = False
  TESTING = False
  #SQLALCHEMY_DATABASE_URI = 'mssql+pymssql://10.147.78.79/asia_assistance_rpa'
  SQLALCHEMY_DATABASE_URI = 'mssql+pymssql://10.147.49.43/asia_assistance_rpa'
  SECRET_KEY = 'RPA_PROD_SECRET'
  #SERVER_NAME = '10.147.78.80:5555'
  SERVER_NAME = '10.147.49.43:5555'

class DemoConfig(BaseConfig):
  ENV = 'demo'
  FLASK_ENV = 'demo'
  DEBUG = True
  TESTING = True
  SQLALCHEMY_DATABASE_URI = 'mssql+pymssql://10.147.49.43/asia_assistance_rpa_demo'
  SECRET_KEY = 'RPA_DEMO_SECRET'
  SERVER_NAME = '10.147.49.43:5560'
