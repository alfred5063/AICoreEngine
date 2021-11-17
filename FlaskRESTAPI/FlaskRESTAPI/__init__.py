#!/usr/bin/python
# Updated as of 1st Dec 2020

import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from os import environ
import sys
from flask_swagger_ui import get_swaggerui_blueprint

db = SQLAlchemy()

def create_app(test_config=None):
  # create and configure the app
  app = Flask(__name__, instance_relative_config=False)

  ### swagger specific ###
  SWAGGER_URL = '/dtiapi'
  API_URL = '/static/swagger.json'
  SWAGGERUI_BLUEPRINT = get_swaggerui_blueprint(
    SWAGGER_URL,
    API_URL,
    config={
        'app_name': "DTI APIs"
    }
  )
  app.register_blueprint(SWAGGERUI_BLUEPRINT, url_prefix=SWAGGER_URL)
  ### end swagger specific ###
 
  env = sys.argv[1] if len(sys.argv) > 1 else 'development'
  config_dict = {
    'local': 'config.LocalConfig',
    'development': 'config.DevelopmentConfig'
  }
  config = config_dict[env]
  app.config.from_object(config)
  try:
    os.makedirs(app.instance_path)
  except OSError:
    pass
  db.init_app(app)


  with app.app_context():
    from FlaskRESTAPI.models import (User, Role, Job, Task, Process_output)
    from FlaskRESTAPI.routes import (register, login, dm_validate_8_functions, cba_msig, automagica_browser, cba_reqdca, finance_ar,
                               mediclinic_mos, finance_updatecml,
                               mediclinic, print_listing, data_migration_client_master, mediclinic_crm, credit_control,
                               data_migration_disbursement_master, data_migration_req, dm_stp,
                               finance_update_max, mediclinic_kpi)
    from FlaskRESTAPI.decorators import before_request, after_request
    
    db.create_all()
    return app
