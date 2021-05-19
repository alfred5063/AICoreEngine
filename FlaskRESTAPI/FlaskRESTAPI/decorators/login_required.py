from functools import wraps
from flaskr.models.User import decode_auth_token, is_valid_user
from flask import request, make_response, abort, g
from flaskr.response.response import Failure

def login_required(f):
  @wraps(f)
  def wrap(*args, **kwargs):
    try:
      if not request.headers.get('Authorization'):
        return Failure('You are not authorize to access this api, kindly to authenticate for request a valid token to access the api.', 401)
      else:
        data = request.headers.get('Authorization')
        token = str.replace(str(data), 'Bearer ', '')
        user_id = decode_auth_token(token)
        user = is_valid_user(user_id)
        if user:
          g.user = user
          return f(*args, **kwargs)
        else:
          return Failure('Your token session expired, please log in again.', 401)
    except Exception as e:
      exception_message = str(e)
      return Failure(exception_message)
  return wrap
