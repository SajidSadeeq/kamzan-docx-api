#docxapi.wsgi
import sys 
sys.path.insert(0, '/var/www/html/docxapi')
  
from docxapi import app as application
