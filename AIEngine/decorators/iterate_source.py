import functools
import os
from datetime import datetime
from directory.get_filename_from_path import get_file_name

def rename_destination(destination, i, file):
  head, tail = get_file_name(destination)
  now = datetime.now().strftime('%d%m%Y%H%M%S')
  timestamp = str(now)
  filename = file.split('.')[0]
  result = head + "\\"+ filename + '_' + timestamp + '.xlsx'
  return result

def iterate_source(func):
  @functools.wraps(func)
  def wrapper_decorator(*args, **kwargs):
    source = kwargs['source']
    if os.path.isdir(source):
      source = kwargs.pop('source', None)
      destination = kwargs.pop('destination', None)
      non_hidden_files = [filename for filename in os.listdir(source) if not filename.startswith('~$') and any(map(filename.endswith, ('.xls', '.xlsx')))]
      result_list = []
      for i, file in enumerate(non_hidden_files):
        new_source = os.path.join(source, file)
        new_destination = rename_destination(destination, i, file) if destination else None
        result = func(*args, source=new_source, destination=new_destination,  **kwargs)
        result_list.append(result)
      return result_list
    else:
      result = func(*args, **kwargs)
      return result
  return wrapper_decorator

