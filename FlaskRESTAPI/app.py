"""
This script runs the application using a development server.
It contains the definition of routes and views for the application.
"""
from os import environ
from FlaskRESTAPI import create_app
import sys
import os
from pathlib import Path

current_path = os.path.dirname(os.path.abspath(__file__))
dti_path = str(Path(current_path).parent) + '\\RPA.DTI'

#This path will use for server environment setup
#sys.path.append(r"C:\Asia-Assistance\RPA.DTI")
sys.path.append(dti_path)
if __name__ == '__main__':
  app = create_app()
  app.run()
