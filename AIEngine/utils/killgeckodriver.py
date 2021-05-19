import os
from pathlib import Path
import psutil

def kill_firefox():
  PROCNAME = "geckodriver" # or chromedriver or IEDriverServer
  for proc in psutil.process_iter():
      # check whether the process name matches
      if proc.name() == PROCNAME:
          proc.kill()

