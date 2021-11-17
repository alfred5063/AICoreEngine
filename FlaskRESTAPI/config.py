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
  SQLALCHEMY_DATABASE_URI = '\\localhost'
  SECRET_KEY = 'RPA_SECRET' # os.urandom(24)
  SERVER_NAME = 'localhost:5555'
  
class LocalConfig(DevelopmentConfig):
  ENV = 'local'
  FLASK_ENV = 'local'
  DEBUG = True
  TESTING = True
  SQLALCHEMY_DATABASE_URI = '\\localhost'
  SERVER_NAME = 'localhost:5555'
