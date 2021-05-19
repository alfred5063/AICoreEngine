import functools

def decorator(func):
  @functools.wraps(func)
  def wrapper_decorator(*args, **kwargs):
    # Do something before.
    return func(*args, **kwargs)
    # Do something after.
  return wrapper_decorator
