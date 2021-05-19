import ntpath

#This return a file name from the full path no matter what os path
def get_file_name(path):
  head, tail = ntpath.split(path)
  return head, tail or ntpath.basename(head)

