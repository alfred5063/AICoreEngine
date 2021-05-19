from flask import current_app as app
from FlaskRESTAPI import db
import datetime
from werkzeug.security import check_password_hash, generate_password_hash
from flask import flash
import jwt

class User(db.Model):
  """Model for user accounts. """
  __tablename__ = 'users'
  id = db.Column(db.Integer, primary_key=True)
  first_name = db.Column(db.String(45), nullable=True)
  last_name = db.Column(db.String(45), nullable=True)
  email = db.Column(db.String(45), unique=True, nullable=False)
  password = db.Column(db.String(100), nullable=False)
  session_id = db.Column(db.Integer, nullable=False)
  created_at = db.Column(db.DateTime, nullable=False)
  # current_process = db.relationship('Process',
  #   primaryjoin='and_(Process.queue_id == Queue.id, Process.status == "in progress")'
  # )

  def __init__(self, first_name=None, last_name=None, email=None, password=None):
    self.first_name = first_name
    self.last_name = last_name
    self.email = email
    self.password = generate_password_hash(password)
    self.created_at = datetime.datetime.now()
    self.session_id = 1

  def __repr__(self):
    return '<User {}'.format(self.email)

def create_user(first_name, last_name, email, password):
  user, error, token = None, None, None
  if not (email and password):
    error = 'email & Password are required.'
  elif user_exist(email):
    error = 'This email: {} is already registered'.format(email)
  if error is None:
    user = User(first_name, last_name, email, password)
    db.session.add(user)
    db.session.commit()
    token = encode_auth_token(user.id)
  return user, error, token

def encode_auth_token(user_id):
  try:
    payload = {
      'exp': datetime.datetime.utcnow() + datetime.timedelta(days=1, seconds=5),
      'iat': datetime.datetime.utcnow(),
      'user_id': user_id
    }
    token = jwt.encode(
      payload,
      app.config.get('SECRET_KEY'),
      algorithm='HS256'
    ).decode('utf-8')
    return token
  except Exception as e:
    return e

def decode_auth_token(auth_token):
  try:
    payload = jwt.decode(auth_token, app.config.get('SECRET_KEY'), algorithms=['HS256'])
    return payload['user_id']
  except jwt.ExpiredSignatureError:
    return 'The session expired, please log in again.'
  except jwt.InvalidTokenError as error:
    print(error)
    return 'Invalid access, please log in again.'

def user_exist(email):
  result = User.query.filter_by(email=email).first()
  return result

def is_valid_user(user_id):
  result = User.query.filter_by(id=user_id).first()
  return result

def request_login(email, password):
  token, error = None, None
  user = User.query.filter_by(
    email=email
  ).first()
  if user and check_password_hash(user.password, password):
    token = encode_auth_token(user.id)
    if token:
      create_new_user_session(user)
      return token, error
  else:
    error = 'Wrong username or password'
    return token, error

def request_login_rpa(email, password):
  token, error = None, None
  user = User.query.filter_by(
    email=email
  ).first()

  if user and user.password == password:
    token = encode_auth_token(user.id)
    if token:
      create_new_user_session(user)
      return token, error
  else:
    error = 'Wrong username or password'
    return token, error

def create_new_user_session(user):
  user.session_id += 1
  db.session.commit()
