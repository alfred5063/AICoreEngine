from directory.get_filename_from_path import get_file_name

class session:
  class base:
    taskid = None
    guid = None
    email = None
    stepname = None
    username = None
    password = None
    filename = None

    def __init__(self, taskid, guid, email, stepname, username=None, password=None, filename=None):
      self.taskid = taskid
      self.guid = guid
      self.email = email
      self.stepname = stepname
      self.username = username
      self.password = password
      self.filename = filename

    def set_password(self, password):
      self.password = password
    def set_filename(self, filename):
      self.filename = filename

  class data_management: 
    source = None
    destination = None
    filename = None
    
    def __init__(self, source, destination=None):
      self.source = source
      self.destination = destination
      head, tail = get_file_name(source)
      self.filename = tail

  class dataframe:
    dataframe = None

    def __init__(self, dataframe):
      self.dataframe = dataframe

    def set_dataframe(dataframe):
      self.dataframe = dataframe
      
  class finance: 
    source = None
    destination = None
    
    def __init__(self, source, destination=None):
      self.source = source
      self.destination = destination
