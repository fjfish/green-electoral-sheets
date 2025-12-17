import sys, os, csv, re
from spreadsheet_functions import *
from datetime import datetime, date
from wirral_config import *

print(sys.argv)

if len(sys.argv) <= 1:
    print("Please area code")
    sys.exit()
    
areas = sys.argv[1:]

for area in areas:
    reorder(area)

