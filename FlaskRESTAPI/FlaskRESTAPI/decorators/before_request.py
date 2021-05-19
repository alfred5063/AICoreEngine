from flask import current_app as app
from flask import g
from flask import request, abort

API_ALLOWED_IPS = ['127.0.0.1']

@app.before_request
def before_request_callback():
  g.user = None
  g.response = None
  g.queue = None
  #limit_remote_addr()

def limit_remote_addr():
  for IP in API_ALLOWED_IPS:
    if str(request.remote_addr).startswith(IP) or str(request.remote_addr) == IP:
      return abort(403)  # Forbidden
