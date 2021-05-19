from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad
import base64
import unicodedata
def string_to_base64(s):
  return base64.b64encode(s.encode('utf-8')).decode('utf-8')

def base64_to_string(b):
  return base64.b64decode(b).decode('utf-8')
def remove_control_characters(s):
    return "".join(ch for ch in s if unicodedata.category(ch)[0]!="C")

def decrypt_pass (password, salt, public_key):
  salt = base64.b64decode(salt) # Salt
  public_key = base64.b64decode(public_key) # Public Key
  password = base64.b64decode(password) #password
  cipher = AES.new(public_key, AES.MODE_CBC, salt)  # Setup cipher
  text = cipher.decrypt(password).decode('utf-8')
  if _unpad(text.strip().replace('\a','').replace('\b','').replace('',''))=="":
    return remove_control_characters(text)
  else:
    return _unpad(text.strip().replace('\a','').replace('\b','').replace('',''))

def _pad(self, s):
  return s + (self.bs - len(s) % self.bs) * chr(self.bs - len(s) % self.bs)

def _unpad(s):
  return s[:-ord(s[len(s)-1:])]

#public_key = bytes(str('RidAeJYPVSNIyPxGlPBlrHeZ6FkWbAcJDkjiQwnveio='), 'ascii')
#salt = bytes(str('HZiFwF5aZlxa70pRolKFLg=='), 'ascii')
#password = 'U0JobmThY1pSYgw4tvaVKQ=='
#decry = decrypt_pass(password, salt, public_key)
#print(decry)
