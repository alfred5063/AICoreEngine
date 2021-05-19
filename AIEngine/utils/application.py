class browser:
  class application:
    appId = None
    appName = None
    appUrl = None

    def __init__(self, appId, appName, appUrl):
      self.appId = appId
      self.appName = appName
      self.appUrl = appUrl

    def set_appUrl(self, appUrl):
      self.appUrl = appUrl
    def set_appName(self, appName):
      self.appName = appName
    def set_appId(self, appId):
      self.appId = appId

  class xpath:
    app_page_key = None
    appId = None
    key = None
    input_xpath = None
    x_coordinate = None
    y_coordinate = None

    def __init__(app_page_key, appId, key, input_xpath, x_coordinate, y_coordinate):
        self.app_page_key = app_page_key
        self.appId = appId
        self.key = key
        self.input_xpath = input_xpath
        self.x_coordinate = x_coordinate
        self.y_coordinate = y_coordinate

    def set_xpath(self, appId, key, input_xpath):
      self.appId = appId
      self.key = key
      self.input_xpath = input_xpath

    def set_xpath_coordinate(self, appId, key, x_coordinate, y_coordinate):
      self.appId = appId
      self.key = key
      self.x_coordinate = x_coordinate
      self.y_coordinate = y_coordinate

 
      
