from flask import current_app as app
from flask import request
import json
from flaskr.models.User import request_login_rpa
from flaskr.response.response import Success, Failure
from flask import render_template
from datetime import datetime
from flaskr.decorators.login_required import login_required

@app.route('/login', methods=['POST'])
def login():
  email = request.form['email']
  password = request.form['password']
  token, error = request_login_rpa(email, password)
  print(token)
  print(error)
  if not error:
    message = Success({
      'message': 'successfully logged in.',
      'auth_token': token
    })
    return message
  return Failure(error)

@app.route('/')
@app.route('/home')
def home():
    """Renders the home page."""
    return render_template(
        'index.html',
        title='Home Page',
        year=datetime.now().year,
    )

@app.errorhandler(403)
def not_found(error):
    return render_template('403.html'), 403

@app.route('/contact')
def contact():
    """Renders the contact page."""
    return render_template(
        'contact.html',
        title='Contact',
        year=datetime.now().year,
        message='Your contact page.'
    )

@app.route('/about')
def about():
    """Renders the about page."""
    return render_template(
        'about.html',
        title='About',
        year=datetime.now().year,
        message='Your application description page.'
    )
