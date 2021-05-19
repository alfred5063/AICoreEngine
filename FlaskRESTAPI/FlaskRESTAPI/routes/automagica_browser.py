from workflow.kill_firefox import automagica_kill_firefox
import json
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required

@app.route('/automagica_kill_firefox_memory', methods = ['POST'])
@login_required
def automagica_kill_firefox():
  try:
    #myjson = json.loads(request.form['json'])
    # Call to kill firefox memory
    automagica_kill_firefox()

    return Success({'message':'Completed.'})
  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )

