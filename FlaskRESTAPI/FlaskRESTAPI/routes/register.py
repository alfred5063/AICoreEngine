from flask import current_app as app
from flask import request
import json
from flaskr.models.User import create_user
from flaskr.response.response import Success, Failure

# a simple page that says hello
@app.route('/register', methods=['POST'])
def register():
  first_name = request.form['first_name']
  last_name = request.form['last_name']
  email = request.form['email']
  password = request.form['password']
  user, error, token = create_user(first_name, last_name, email, password)
  if not error:
    message = Success({
      'status': 'sucess',
      'message': 'Successfully registered.',
      'auth_token': token,
    })
    return message
  return Failure(error)
