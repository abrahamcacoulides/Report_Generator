import sys
import os
import django
import datetime
import openpyxl
from openpyxl.styles import PatternFill, NamedStyle
from Variables import *

#*****************************
sys.path.append(project_location)#folder for the app
#*****************************
os.environ['DJANGO_SETTINGS_MODULE'] = 'testwebpage.settings'
django.setup()

from prodfloor.models import Info,Times,Stops
from extra_functions import *

#Change this values if first setup
#*****************************
u_input = input('Please insert the date(mm/dd/yyyy):')
dates_list = u_input.split('/')
user = user_var
day = datetime.date(int(dates_list[2]),int(dates_list[0]),int(dates_list[1]))
folder_name = ''
#*****************************