#!/usr/bin/python
# FINAL SCRIPT updated as of 12th NOVEMEBR 2019
# Workflow - CBA/print listing
# Version 1
from flask import current_app as app
from flask import request,jsonify
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
from workflow.workflow_print_listing import process_move_files_to_archieved,process_get_print_listing_archieved,process_get_print_listing_new
import json

@app.route('/get_print_listing_new', methods = ['GET'])
@login_required
def get_print_listing_new():
  try:
    data = process_get_print_listing_new()
    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )

@app.route('/move_files_to_archieved', methods = ['POST'])
@login_required
def move_files_to_archieved():
  try:
    files_name=request.form['files_name']
    email=request.form['email']
    data = process_move_files_to_archieved(files_name,email)
    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )


@app.route('/get_print_listing_archieved', methods = ['GET'])
@login_required
def get_print_listing_archieved():
  try:
    data = process_get_print_listing_archieved()
    return Success({'message':'Completed.', 'data': data})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )

