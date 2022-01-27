#docxapi.wsgi
import sys 
sys.path.insert(0, '/var/www/docxapi')
  
from docxapi import app as application
