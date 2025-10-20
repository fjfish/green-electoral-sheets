import sys, os, csv, re
from spreadsheet_functions import *
from datetime import datetime, date
from wirral_config import *

if len(sys.argv) <= 3:
    print("Please specify election name, yearcode of used register and a list of files to import")
    sys.exit()
    
election = sys.argv[1]
year = sys.argv[2]

filenames = []
for arg in sys.argv[3:]:
    if os.path.isfile(arg):
        filenames.append(arg)
    else:
        print("could not find %s" % arg)
        
NO_SUFFIX_REGEX = re.compile(r"^(.{2})-(\d+)$")
SUFFIX_REGEX = re.compile(r"^(.{2})-(\d+)/(\d+)$")

areas = set()
results = {}
def add(code, electorNumber):
    areas.add(code[0])
    if code not in results:
        results[code] = set()
        postal_voters[code] = set()
    results[code].add(electorNumber)

postal_voters = {}
def add_postal(code, electorNumber):
    areas.add(code[0])
    if code not in postal_voters:
        results[code] = set()
        postal_voters[code] = set()
    postal_voters[code].add(electorNumber)

class EmptyIDException(Exception):
    pass
            
def getCode(year, row):
    if not row["ElectorNo"]:
         raise EmptyIDException()
    no_suffix_match = NO_SUFFIX_REGEX.match(row["ElectorNo"])
    if no_suffix_match:
        code, eid = no_suffix_match.groups()
        return code[0], "%s:%s-%04i" % (year, code, int(eid))
    else:
        suffix_match = SUFFIX_REGEX.match(row["ElectorNo"])
        if suffix_match:
            code, eid, sufix = suffix_match.groups()
            return code[0], "%s:%s-%04i/%02i" % (year, code, int(eid), int(sufix))
        else:
            raise Exception("Could not match %s." % row["ElectorNo"])
            


for filename in filenames:
    with open(filename, encoding='latin1') as resultsCSV:
        Ereader = csv.DictReader(resultsCSV, delimiter=',')
        for row in Ereader:
            try:
                area, fullCode = getCode(year, row)
                if row["PVSStatus"] not in [0, "0", ""]:
                    add(area, fullCode)
                if row["Type"] in ["Postal", "PostalProxy"]:
                    add_postal(area, fullCode)
            except EmptyIDException:
                pass

for area in areas: 
  if area == "D":
    try:
        add_marks(area, results[area], election, postal=postal_voters[area])
    except UnknownAreaException as e:
        print(str(e))
        pass
