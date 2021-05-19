from functools import wraps
from flaskr.models.User import decode_auth_token, is_valid_user
from flask import request, make_response, abort, g
from flaskr import db
from flaskr.models.Role import Role
from flaskr.response.response import Failure
from flaskr.models.User import User
# from sqlalchemy import select
from sqlalchemy.orm import Load

def roles_required(*roles, operation='or'):
  def wrapper(f):
    @wraps(f)
    def wrapped(*args, **kwargs):
      action = request.path
      user_actions = get_current_user_actions()
      if is_valid_action(user_actions, action):
        return f(*args, **kwargs)
      return Failure('Permission denined', 403)
    return wrapped
  return wrapper

def get_current_user_role():
  roles = g.user.roles
  user = g.user
  roles_names = [role.name for role in roles]
  return roles_names

def get_current_user_actions():
  user = g.user
  print(user.id)
  actions = db.session.query(Action).select_from(User).join(user_roles).filter(User.id==user.id).join(Role).join(role_actions).join(Action).all()
  action_names = [action.name for action in actions]
  return action_names

def is_valid_action(user_actions, action):
  print(user_actions)
  print(action)
  return action in user_actions

def is_valid_roles(user_roles, roles):
  return set(user_roles).intersection(set(roles))
  # return set(user_roles) >= set(roles) # for 'and' operation
