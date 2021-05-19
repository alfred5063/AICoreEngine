import json
from flask import current_app as app
from flask import request, g
from flaskr.response.response import Response

@app.after_request
def after_request_callback(res):
  construct_log(res)
  return res

def is_logged_in():
  return g.user.id if g.user else Response.ANONYMOUS

def get_session_id():
  if g.queue:
    return g.queue.id
  return g.user.session_id if g.user else None

def construct_log(res):
  action_name = '{} {}'.format(request.method, request.url)
  #parameter_string = json.dumps({
  #  'args': json.dumps(request.args),
  #  'form': json.dumps(request.form)
  #})
  #user_id = is_logged_in()
  message = json.dumps(g.response)
  print(g.response)
  #status = g.response['status']
  #session_id = get_session_id()

def after_request(response):
  response.headers.add('Access-Control-Allow-Origin','*')
  response.headers.add('Access-Control-Allow-Headers', "Authorization, Content-Type")
  response.headers.add('Access-Control-Expose-Headers', "Authorization")
  response.headers.add('Access-Control-Allow-Methods', "GET, POST, PUT, DELETE, OPTIONS")
  response.headers.add('Access-Control-Allow-Credentials', "true")
  response.headers.add('Access-Control-Max-Age', 60 * 60 * 24 * 20)
  return response
