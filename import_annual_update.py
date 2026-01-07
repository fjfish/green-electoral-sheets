import sys, os, csv, re

from spreadsheet_functions import *
from datetime import datetime, date

from wirral_config import *
from thefuzz import fuzz

DEBUG = True



if len(sys.argv) < 3:
    print("Please specify the file of new data, yearcode, added date and optionally a list of wards to update.")
    sys.exit()
    
csv_filename = sys.argv[1]
year = sys.argv[2]
added_date = sys.argv[3]

if len(sys.argv) > 4:
    wards = []
    for ward in sys.argv[4:]:
        if len(ward) == 1:
            wards.append(ward)
        else:
            print("wards should be represented by single characters not %s" % ward)
else:
    wards = SHEETS.keys()
    
client = get_client()

class electorate(object):
    def __init__(self, d):
        self.d = {}
        for key, value in d.items():
            if type(value) is str:
                self.d[key] = value.replace("\xa0", " ")
            else:
                self.d[key] = value
        self.address = "@".join([str(self.d["Address%i" % i]).strip().lower() for i in range(1,7)])
        if "Elector Surname Elector Forename" in d.keys():
            self.d["Elector Name"] = self.d["Elector Surname Elector Forename"]
        if "Elector Name" in d.keys():
            names = self.d["Elector Name"].strip().split(" ", 2)
            names = names + ["" for i in range(max(0, 3 - len(names)))]
            self.forename = names[1].lower().strip()
            self.middle_initials = names[2].lower().strip()
            self.surname = names[0].lower().strip()
        else:
            forename_initials = self.d["Elector Forename"].strip().split(" ", 1)
            self.forename = forename_initials[0].lower().strip()
            if len(forename_initials) > 1:
                self.middle_initials = forename_initials[1].lower().strip()
            else:
                self.middle_initials = ""
            self.surname = self.d["Elector Surname"].lower().strip()
            self.d["Elector Name"] = self.d["Elector Forename"] + " " + self.d["Elector Surname"]
    
new_electorate = []
address_name_lookup = {}
with open(csv_filename) as newcsvfile:
    new_reader = csv.DictReader(newcsvfile, delimiter=',')
    for row in new_reader:
        if row["Elector DOB"] != "":
            row["Elector DOB"] = datetime.strptime(row["Elector DOB"], "%d-%b-%y").strftime("%Y-%m-%d")
        row["Postcode"] = row["PostCode"]
        e = electorate(dict(row))
        new_electorate.append(e)
        if e.address not in address_name_lookup:
            address_name_lookup[e.address] = []
        address_name_lookup[e.address].append(e.d["Elector Name"])

comparisons = [lambda newID, oldID, full_match_address: 
                 newID.address == oldID.address and 
                 newID.forename == oldID.forename and 
                 newID.middle_initials == oldID.middle_initials and 
                 newID.surname == oldID.surname,
             lambda newID, oldID, full_match_address: 
                 newID.forename == oldID.forename and 
                 newID.middle_initials == oldID.middle_initials and 
                 newID.surname == oldID.surname,
             lambda newID, oldID, full_match_address: 
                 newID.address == oldID.address and 
                 newID.forename == oldID.forename and 
                 newID.surname == oldID.surname,
             lambda newID, oldID, full_match_address: 
                 newID.address == oldID.address and 
                 newID.middle_initials == oldID.middle_initials and 
                 newID.forename == oldID.forename and
                 newID.address in full_match_address,
             lambda newID, oldID, full_match_address: 
                 newID.address == oldID.address and 
                 newID.forename == oldID.forename and
                 newID.address in full_match_address,
             lambda newID, oldID, full_match_address: 
                 newID.address == oldID.address and 
                 newID.forename == oldID.surname and
                 newID.surname == oldID.forename,
             lambda newID, oldID, full_match_address: 
                 newID.address == oldID.address and 
                 fuzz.ratio(newID.d["Elector Name"], oldID.d["Elector Name"]) > 90,
             lambda newID, oldID, full_match_address: 
                 newID.address == oldID.address and 
                 fuzz.ratio(newID.d["Elector Name"], oldID.d["Elector Name"]) > 80
                 ]
print(wards)
for area in wards:
    print(area)
    sheets, masterSheet, headings, records = read_sheet(area)
    old_electorate = []
    for old_result in records:
        old_electorate.append(electorate(old_result))
    
    copy_keys = headings.copy()
    for unused_keys in ["Elector Number", "Elector Markers", "Elector DOB", "Elector Name"] + \
                       ["Address1", "Address2", "Address3", "Address4", "Address5", "Address6"] + \
                       ["Postcode", "House", "Road", "Round", "LE Knock Priority", "GE Knock Priority"]:
        copy_keys.remove(unused_keys)

    new_matched = []
    full_match_address = set()
    for i, test in enumerate(comparisons):
        new_electorate2 = []
        for new in new_electorate:
            for old in old_electorate:
                if test(new, old, full_match_address):
                    for key in copy_keys:
                        new.d[key] = old.d[key]
                    
                    new.d["Previous Elector Number"] = old.d["Elector Number"]
                    #new.d["Match Method"] = i
                    #new.d["old_address_1"] = old.d["Address1"]
                    #new.d["old_address_2"] = old.d["Address2"]
                    #new.d["old_address_3"] = old.d["Address3"]
                    #new.d["old_address_4"] = old.d["Address4"]
                    #new.d["old_address_5"] = old.d["Address5"]
                    #new.d["old_address_6"] = old.d["Address6"]
                    #new.d["old_name"] = old.d["Elector Name"]
                    if i == 0:
                        full_match_address.add(new.address)
                    new_matched.append(new)
                    old_electorate.remove(old)
                    break
            else:
                new_electorate2.append(new)
        print("Method %i got %i matches" % (i, len(new_electorate) - len(new_electorate2)))
        new_electorate = new_electorate2
    print("%i new electors." % len(new_electorate))
    print("%i deleted electors." % len(old_electorate))
    #Handle unmatched
    for new in new_electorate:
        for key in copy_keys:
            new.d[key] = ""
        new.d["Added Date"] = added_date
        new.d["Match Method"] = -1
        new_matched.append(new)
    
    #Write data back
    if not DEBUG:
        new_records = [n.d for n in new_matched]
        for record in new_records:
            make_electoral_number(record, year)
        
        deleted_records = [n.d for n in old_electorate]
        
        write_deleted_record(area, headings, deleted_records, added_date)
        
        write_record(area, headings, new_records)
