#!/usr/bin/python
# FINAL SCRIPT updated as of 29th April 2020
# Data Migration Disbursement Master

# Declare Python libraries needed for this script
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_process_data_migration_dm import process_data_migration_DM_incremental_load

@app.route('/process_data_migration_dm', methods = ['POST'])
@login_required
def process_data_migration_dm():
  try:

    migration_folder = request.form['migration_folder']
    year = request.form['year']
    email = request.form['email']

		# Call out main function from process_data_migration_CM.py
    data = process_data_migration_DM_incremental_load(migration_folder, year, email)

    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )

