
from datetime import datetime, date
from spreadsheet_functions import *
from wirral_config import *

area = "Q"

sheet_id, sheet_name = SHEETS[area]

client = get_client()

input_sheet_id = "1bEBna8fNHNZyJMV8I5DmjMIikRe2Fxlfv5cOarTYXtA"
input_sheet = client.open_by_key(input_sheet_id).worksheet("Q_Dec2024")
record_list = input_sheet.get_all_records()

headings = ["First Name", "House", "Road", "Round", "Contacted", "LE Rating", "Party", "Mobile", "Consent - Mobile", "Date - Consent - Mobile", "Recorded By - Consent - Mobile", "Consent - WhatsApp", "Date - Consent - WhatsApp", "Recorded By - Consent - WhatsApp", "Landline", "Consent - Landline", "Date - Consent - Landline", "Recorded By - Consent - Landline", "Email", "Consent - Email", "Date - Consent - Email", "Recorded By - Consent - Email", "Poster", "Board", "DNC", "DND", "Notes", "GE Rating", "Chosen Name", "Full Name", "Elector Number", "MM24", "GE24", "Address Type", "PV", "Added Date", "Elector Name", "Elector Markers", "Elector DOB", "Previous Elector Number", "Address1", "Address2", "Address3", "Address4", "Address5", "Address6", "Postcode", "AddressFull", "Last Contact", "DNC Date", "DND Date", "Poster Request Date", "Board Request Date"]
           


def calculate_age(born, on = date.today()):
    return on.year - born.year - ((on.month, on.day) < (born.month, born.day))

for record in record_list:
    make_electoral_number(record, "24-25")
    record["Postcode"] = record["PostCode"].strip().replace("\xa0", " ")
    record["Added Date"] = ""
    record["Previous Elector Number"] = ""
    record["MM24"] = ""
    record["GE24"] = ""
    record["PV"] = ""
    record["Board"] = "FALSE"
    record["Board Request Date"] = ""
    record["Email"] = ""
    record["Landline"] = ""
    record["Mobile"] = ""
    record["Poster"] = "FALSE"
    record['Poster Request Date'] = ""
    record["LE Rating"] = ""
    record["GE Rating"] = ""
    record["Notes"] = ""
    
    ns=record["Elector Surname Elector Forename"].split(" ")
    record["Elector Name"] = " ".join([ns[-1]] + ns[:-1])
    
    record["DNC"] = "FALSE"
    record["DNC Date"] = ""

    if record["Elector DOB"] != "":
        record["Elector DOB"] = datetime.strptime(record["Elector DOB"], "%d-%b-%y").strftime("%Y-%m-%d")
    else:
        record["Elector DOB"] = ""

    record["Party"] =  ""
    record['Chosen Name'] = ""
    record["Last Contact"] = ""

create_data_for_new_cols(record_list)    
write_record(area, headings, record_list)
write_addresses(area, record_list)
