import sys, os, csv, re
from spreadsheet_functions import *
from datetime import datetime, date
from wirral_config import *

if len(sys.argv) <= 3:
    print("""Please specify election name, yearcode of used register and a list of files to import
Eg. import_marked.py GE24 23-24 voters.csv""")
    sys.exit()
    
election = sys.argv[1]
year = sys.argv[2]

filenames = []
for arg in sys.argv[3:]:
    if os.path.isfile(arg):
        filenames.append(arg)
    else:
        print("could not find %s" % arg)
        
NO_SUFFIX_REGEX = re.compile(r"^(.{2})-?(\d+)$")
SUFFIX_REGEX = re.compile(r"^(.{2})-?(\d+)/(\d+)$")

results = {}
def add(code, electorNumber):
    if code not in results:
        results[code] = set()
    results[code].add(electorNumber)

for filename in filenames:
    with open(filename) as resultsCSV:
        GEreader = csv.reader(resultsCSV, delimiter=',')
        for a, b, elector_id in GEreader:
            no_suffix_match = NO_SUFFIX_REGEX.match(elector_id)
            if no_suffix_match:
                code, eid = no_suffix_match.groups()
                add(code[0], "%s:%s-%04i" % (year, code.upper(), int(eid)))
            else:
                suffix_match = SUFFIX_REGEX.match(elector_id)
                if suffix_match:
                    code, eid, sufix = suffix_match.groups()
                    add(code[0], "%s:%s-%04i/%02i" % (year, code.upper(), int(eid), int(sufix)))
                else:
                    print("Could not match %s." % elector_id)


for area, elector_ids in results.items():
    add_marks(area.upper(), elector_ids, election)

