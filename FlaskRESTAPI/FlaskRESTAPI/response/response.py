import json
from flask import make_response, g
import linecache
import sys

def print_exception():
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    print('EXCEPTION IN ({}, LINE {} "{}"): {}'.format(filename, lineno, line.strip(), exc_obj))

def pack_response(dictionary):
  # return json as a string
  return json.dumps(dictionary)

def is_dict(obj):
  return isinstance(obj, dict)

class Response():
  SUCCESS = 'Success' # Static variable
  FAIL = 'Fail' # Static variable
  ANONYMOUS = 'Anonymous' # Static variable

  def __new__(cls, message, status, response_code):
    response_dict = None
    if is_dict(message):
      response_dict = {
        'message': message,
        'status': status
      }
      response_dict = pack_response(response_dict)
    else:
      response_dict = pack_response({
        'message': message,
        'status': status
      })
    response = make_response(response_dict, response_code)
    g.response = json.loads(response_dict)
    return response

class Success():
  def __new__(cls, message, response_code=200):
    return Response(message, Response.SUCCESS, response_code)

class Failure():
  def __new__(cls, message, response_code=500, debug=False):
    if debug:
      print_exception()
    return Response(message, Response.FAIL, response_code)
