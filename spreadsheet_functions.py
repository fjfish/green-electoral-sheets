from datetime import datetime, date
from wirral_config import *
from google.oauth2.service_account import Credentials
import gspread
from gspread.utils import ValidationConditionType
import re
from random import random
from time import sleep
from gspread_formatting import get_data_validation_rule, set_data_validation_for_cell_range
from datetime import date
from collections import OrderedDict

COLUMN_WIDTHS = {"Given Name": 80,
                 "Full Name": 150,
                 "Chosen Name": 150,
                 "Elector Name": 150,
                 "House": 100,
                 "Road": 100,
                 "Round": 50,
                 "Elector Number": 100,
                 "Elector Markers": 50,
                 "Elector DOB": 80,
                 "Previous Elector Number": 100,
                 "Address1": 150,
                 "Address2": 150,
                 "Address3": 150,
                 "Address4": 150,
                 "Address5": 150,
                 "Address6": 150,
                 "Postcode": 90,
                 "AddressFull": 100,
                 "Added Date": 80,
                 "PV": 70,
                 "LE Knock Priority": 120,
                 "LE Rating": 80,
                 "Party": 80,
                 "Mobile": 100,
                 "Landline": 100,
                 "Email": 100,
                 "Poster": 70,
                 "Board": 70,
                 "DNC": 70,
                 "Notes": 200,
                 "GE Rating": 80,
                 "GE Knock Priority": 120,
                 "Poster Request Date": 100,
                 "Board Request Date": 100,
                 "Contacted": 70,
                 "Address Type": 80,
                 "Consent - Mobile": 80,
                 "Date - Consent - Mobile": 60,
                 "Recorded By - Consent - Mobile": 105,
                 "Consent - WhatsApp": 155,
                 "Date - Consent - WhatsApp": 60,
                 "Recorded By - Consent - WhatsApp": 105,
                 "Consent - Landline": 80,
                 "Date - Consent - Landline": 60,
                 "Recorded By - Consent - Landline": 105,
                 "Consent - Email": 80,
                 "Date - Consent - Email": 60,
                 "Recorded By - Consent - Email": 105}
COLUMN_WIDTH_DEFAULT = 60    

DATE_HEADERS = ["Elector DOB", "Added Date", "Poster Request Date", "Board Request Date", "DNC Date", "DND Date", "Date - Consent - Email", "Date - Consent - Mobile", "Date - Consent - WhatsApp", "Date - Consent - Landline"]

BOOL = {"TRUE": True, "FALSE": False, True: True, False: False, "": False}

MEDIUMS_RECORDS =  [("Email", "Email", "=EmailConsentTemplates"), 
                    ("Landline", "Landline", "=LandlineConsentTemplates"), 
                    ("Mobile", "Mobile", "=MobileConsentTemplates"), 
                    ("WhatsApp", "Mobile", "=WhatsAppConsentTemplates")]

def get_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets"
    ]
    creds = Credentials.from_service_account_file("wirral-green-party-data-f15572884310.json", scopes = scopes)
    return gspread.authorize(creds)

def get_column_letter(col_idx):
    """Convert a column number into a column letter (3 -> 'C')

    Right shift the column col_idx by 26 to find column letters in reverse
    order.  These numbers are 1-based, and can be converted to ASCII
    ordinals by adding 64.

    """
    # these indicies corrospond to A -> ZZZ and include all allowed
    # columns
    if not 1 <= col_idx <= 18278:
        raise ValueError("Invalid column index {0}".format(col_idx))
    letters = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx, 26)
        # check for exact division and borrow if needed
        if remainder == 0:
            remainder = 26
            col_idx -= 1
        letters.append(chr(remainder+64))
    return ''.join(reversed(letters))

def calculate_age(born, on = date.today()):
    return on.year - born.year - ((on.month, on.day) < (born.month, born.day))

THISYEAR = date.today().year - 2000
FIVEYEARSAGO = THISYEAR - 5
TENYEARSAGO = THISYEAR - 10

ELECTION_REGEX = re.compile(r"^(GE|LE|EU|MM|PCC)(\d\d)$")

r"""
def knock_priority(record, last_election_date, last_election_code, recent_elections, last_election_codes):
    if record["Elector DOB"] != "":
        dob = datetime.fromisoformat(record["Elector DOB"].strip())
    else:
        dob = None
    if dob and calculate_age(dob) < 18:
        return 9999
    if last_election_code in record and BOOL[record[last_election_code]]:
        return 1
    if record["Added Date"] and datetime.strptime(record["Added Date"].strip(), "%Y-%m-%d") > last_election_date:
        return 2
    if dob and calculate_age(dob) >= 18 and calculate_age(dob, last_election_date) < 18:
        return 3
    if any([election_code in record and BOOL[record[election_code]] for election_code in recent_elections]):
        return 4
    election_matches = [(election, re.match(r"^(%s)(\d\d)$" % "|".join(last_election_codes), election)) for election in record]
    five_year_election_matches = [election for (election, match) in election_matches if match and int(match.group(2))>FIVEYEARSAGO]
    if any([election in record and BOOL[record[election]] for election in five_year_election_matches]):
        return 5
    ten_year_election_matches = [election for (election, match) in election_matches if match and int(match.group(2))>TENYEARSAGO]
    if any([BOOL[record[election]] for election in ten_year_election_matches]):
        return 6
    return 7
    
def determine_knock_priority(record, area):
        LE = LOCAL_ELECTION[area]
        record["LE Knock Priority"] = knock_priority(record, LE["last"], LE["name"], LE["recent"], ["GE", "LE", "MM", "PCC"])
        GE = GENERAL_ELECTION[area]
        record["GE Knock Priority"] = knock_priority(record, GE["last"], GE["name"], [GE["name"]], ["GE"])
"""
  
def make_electoral_number(record, year):
    record["Elector Number"] = "%s:%s-%04i" % (year, record["Elector Number Prefix"].strip(), int(record["Elector Number"]) )
    if record["Elector Number Suffix"] not in [0, "0", ""]:
        record["Elector Number"] += "/%02i" % int(record["Elector Number Suffix"])

def make_House_Road_AddressFull(record, area):
    addresses = [str(record["Address%i" % i]).strip().replace("\xa0", " ") for i in range(1,7)]
    if addresses[0] == "":
        record["House"] = ""
        record["AddressFull"] = ""
        record["Road"] = ""
        return
    if addresses[-1] in addresses[:-1]:
        record["Address6"] = ""
        addresses = addresses[:-1]
    record["AddressFull"] = ", ".join([str(record["Address%i" % i]).strip().replace("\xa0", " ") for i in range(1,6) if str(record["Address%i" % i]).strip() != ""])
    postcode_index = addresses.index(record["Postcode"].strip() )
    if postcode_index == 2:
        road_record = addresses[1] # One strange address in Birkenhead
    else:
        road_record = addresses[postcode_index - 3]

    if road_record[0].isdigit():
        house_number, road = road_record.split(" ", 1)
        if house_number.endswith("-") and road[0].isdigit(): # Deal with 4- 6 Knowsley Road
            n, road = road.split(" ", 1)
            house_number = house_number + n
        record["Road"] = road
        house_number = [house_number]
    else:
        for road in ROADS[area]:
            if road_record.endswith(" " + road):
                record["Road"] = road
                house_number = [road_record[:-len(road) + 1]]
                break
        else:
            record["Road"] = road_record
            house_number = []
    record["House"] = ", ".join(addresses[:postcode_index - 3] + house_number)
    
def write_record(area, headings, record_list):
    HOUSE_COLUMN = get_column_letter(headings.index("House") + 1)
    ROAD_COLUMN = get_column_letter(headings.index("Road") + 1)
    CHOSEN_NAME_COLUMN = get_column_letter(headings.index("Chosen Name") + 1)
    ELECTOR_NAME_COLUMN = get_column_letter(headings.index("Elector Name") + 1)
    FULL_NAME_COLUMN = get_column_letter(headings.index("Full Name") + 1)
    ADDRESS_FULL_COLUMN = get_column_letter(headings.index("AddressFull") + 1)
    for n, record in enumerate(record_list):
        make_House_Road_AddressFull(record, area)
        #determine_knock_priority(record, area)
        row = n + 2
        #record["Round"] = f'=if(isna(match({ROAD_COLUMN}{row},RoadsInLookup,0)),' + \
        #                  f'vlookup(concatenate({HOUSE_COLUMN}{row}, " ", {ROAD_COLUMN}{row}),HouseLookup,2,false), ' + \
        #                  f'vlookup({ROAD_COLUMN}{row},RoadLookup,2,false))'
        record["Round"] = f'=if(isna(match({ROAD_COLUMN}{row},PartialRoads,0)),' + \
                          f'vlookup({ROAD_COLUMN}{row}, RoadLookup, 2, false), ' + \
                          f'vlookup(concatenate({HOUSE_COLUMN}{row}, " ", {ROAD_COLUMN}{row}),HouseLookup,2,false)) '
        
        record["Full Name"] = f'=if({CHOSEN_NAME_COLUMN}{row}="", concatenate(IFNA(concatenate(REGEXEXTRACT({ELECTOR_NAME_COLUMN}{row}, "^\\s*(?:\\S+\\s+)?(.*?)\\s*$"), " "), ""), REGEXEXTRACT({ELECTOR_NAME_COLUMN}{row}, "^\\s*(\\S+)")), {CHOSEN_NAME_COLUMN}{row})'
        record["First Name"] = f'=REGEXEXTRACT({FULL_NAME_COLUMN}{row}, "^\\s*(?:Md\\s+)?(?:-\\s+)?(\\S+)")'
        record["Address Type"] = f'''=IfNA(vlookup({ADDRESS_FULL_COLUMN}{row}, 'Address Types'!A:B, 2, False), "")'''


    sheet_id, sheet_name = SHEETS[area]
    sheets = get_client().open_by_key(sheet_id)
    oldsheet = sheets.worksheet(sheet_name)
    oldsheet.update_title(f"Old{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(len(record_list) + 1)
    sheet = sheets.add_worksheet(title="Master", rows=len(record_list) + 1, cols=len(headings),index=0)
    sheet.update([headings], 'A1')
    sheet.update([[record[h] for h in headings] for record in record_list], 'A2', raw=False)
    update_sheet = sheets.worksheet("Update")
    set_formating(sheet, headings, update_sheet, sheets)
    
    #update_sheet.update([[10]], "A1")


def set_formating(sheet, headings, update_sheet, sheets):
    sheet.freeze(1,4)
    formats = []
    formats.append({"range": 'A1:1', "format": {'textFormat': {'bold': True}, "horizontalAlignment" : "LEFT", "backgroundColor": {"red": 0.8,  "green": 1,  "blue": 0.8,  "alpha": 1}}})
    col_widths = []
    unprotected_cols = []
    headingLetterLookup = dict([(heading, colindex) for colindex, heading in enumerate(headings, 1)])
    for colindex, heading in enumerate(headings, 1):
        column = get_column_letter(colindex)
        if heading in ["Poster", "Board", "DNC", "DND", "PV", "Nursing", "Flats", "Contacted"] or ELECTION_REGEX.match(heading):
            exponential_backoff(sheet.add_validation,
                                f'{column}2:{column}',
                                ValidationConditionType.boolean,
                                [],
                                strict=False,
                                )
        elif heading in ["Party"]:
            exponential_backoff(copy_validation_to_col, sheets, update_sheet, 1, 1, sheet, colindex)
            formats.append({"range": f'{column}2:{column}', "format": {"numberFormat": {"type": "TEXT"}, "horizontalAlignment" : "LEFT"}})
        elif heading in ["LE Rating", "GE Rating"]:
            exponential_backoff(copy_validation_to_col, sheets, update_sheet, 2, 1, sheet, colindex)
            formats.append({"range": f'{column}2:{column}', "format": {"numberFormat": {"type": "TEXT"}, "horizontalAlignment" : "LEFT"}})
        elif heading in DATE_HEADERS:
            formats.append({"range": f'{column}2:{column}', "format": {"numberFormat": {"type": "DATE"}, "horizontalAlignment" : "LEFT"}})
        elif heading in ["Email"]:
            sheet.add_validation(f'{column}2:{column}', ValidationConditionType.text_is_email, [],
                                strict=True,
                                inputMessage='Invalid Email address entered.',
                                )
            formats.append({"range": f'{column}2:{column}', "format": {"numberFormat": {"type": "TEXT"}, "horizontalAlignment" : "LEFT"}})
        else:
            formats.append({"range": f'{column}2:{column}', "format": {"numberFormat": {"type": "TEXT"}, "horizontalAlignment" : "LEFT"}})
        for comm_medium, comm_record, options_range in MEDIUMS_RECORDS:
            consent_prefix = "Consent - "
            if heading == f"{consent_prefix}{comm_medium}":
                exponential_backoff(consent_formatting, sheets, sheet, colindex, headingLetterLookup[comm_record])
                exponential_backoff(setDropDownFromRange, sheets, sheet, colindex, options_range)
            if heading == f"Date - Consent - {comm_medium}":
                exponential_backoff(consent_formatting, sheets, sheet, colindex, headingLetterLookup[consent_prefix + comm_medium])
            if heading == f"Recorded By - Consent - {comm_medium}":
                exponential_backoff(consent_formatting, sheets, sheet, colindex, headingLetterLookup[consent_prefix + comm_medium])
            
        if heading in ["First Name", "Full Name", "Elector Number", "Elector Markers", "Elector DOB", "Elector Name", "Previous Elector Number"] + \
                      ["Address%i" % i for i in range(1,7)] + ["Postcode", "AddressFull", "House", "Road", "Round", "Last Contact", "Address Type"] + \
                      ["DNC Date", "DND Date", "Poster Request Date", "Board Request Date"] +\
                      ["Added Date", "PV", "LE Knock Priority", "GE Knock Priority"] or ELECTION_REGEX.match(heading):
            formats.append({"range": f'{column}2:{column}', "format": {"backgroundColor": {"red": 0.8,  "green": 1,  "blue": 0.8,  "alpha": 1}}})
        else:
            unprotected_cols.append(colindex)
        if heading in ["Poster Request Date", "Board Request Date", "Last Contact", "DNC Date", "DND Date", "Poster Request Date", "Board Request Date"]:
            unprotected_cols.append(colindex)
        if heading in COLUMN_WIDTHS:
            col_widths.append((colindex, COLUMN_WIDTHS[heading]))
        else:
            col_widths.append((colindex, COLUMN_WIDTH_DEFAULT))
            
    exponential_backoff(sheet.batch_format, formats)
    exponential_backoff(batch_colwidths, sheets, sheet, col_widths) 
    exponential_backoff(protectSheet, sheets, sheet, unprotected_cols)   
    #Groups do not work if a cell in read only
    """groupCols = [colindex 
                    for colindex, heading 
                    in enumerate(headings, 1) 
                    if heading in ["Full Name", "Elector Number", "Elector Name", "Elector Markers", "Elector DOB", "Previous Elector Number"] + \
                                   ["Address%i" % i for i in range(1,7)] + \
                                   ["Postcode", "AddressFull", "Added Date", "PV", "LE Knock Priority", "GE Knock Priority", "Poster Request Date", "Board Request Date"]
                      or ELECTION_REGEX.match(heading)
                    ]
    for groupCol in groupCols:
         exponential_backoff(sheet.add_dimension_group_columns, groupCol - 1, groupCol)"""
    
    exponential_backoff(sheet.set_basic_filter, f"A1:{get_column_letter(sheet.col_count)}{sheet.row_count}")
     
    name_ranges = [("Master_Headers", get_header_range(sheet)),
                   ("Master_Data", get_data_range(sheet))]
    for i in ["Full Name", "Road", "DNC", "DND", "Added Date", "AddressFull", "Round", "LE Rating", "GE Rating", "Old Round", "PV", "Elector DOB", "Last Contact", "Address Type", "Party", "Mobile", "Consent - Mobile", "Consent - WhatsApp", "Landline", "Consent - Landline", "Email", "Consent - Email"]:
        if i in headingLetterLookup:
            name = "Master_" + i.replace(" ", "_").replace("(", "_").replace(")", "_").replace("-", "_")
            name_ranges.append((name, get_col_data_range(sheet, headingLetterLookup[i])))
    start_election_col, end_election_col = find_election_cols(headings)
    name_ranges.append(("Master_Election_Names", get_header_range(sheet, start_election_col, end_election_col)))
    name_ranges.append(("Master_Elections", get_data_range(sheet, start_election_col, end_election_col)))
    name_ranges.append(("Master_House_Road", get_data_range(sheet, headingLetterLookup["House"]-1, headingLetterLookup["House"] + 1)))
    current_name_ranges = sheets.list_named_ranges()
    set_named_ranges_with_errors(sheets, 
                                 name_ranges, 
                                 dict([(x["name"], x["namedRangeId"]) 
                                       for x 
                                       in current_name_ranges
                                       ])
                                 )

def find_election_cols(headings):
    start_election_col = None
    for colindex, heading in enumerate(headings, 1):
        if start_election_col is None and re.match(r"^(?:LE|GE|MM)\d\d$", heading):
            start_election_col = colindex - 1
        if start_election_col is not None and not re.match(r"^(?:LE|GE|MM)\d\d$", heading):
            end_election_col = colindex - 1
            break
    return start_election_col, end_election_col
   
error_colour = {"red": 250/255, "green": 80/255, "blue": 83/255}
warning_colour = {"red": 255/255, "green": 222/255, "blue": 33/255}



def get_col_data_range(sheet, col_index):
    return  {
              "sheetId": sheet.id,
              "startColumnIndex": col_index - 1,
              "endColumnIndex": col_index,
              "startRowIndex": 1,
              "endRowIndex": sheet.row_count,
            }
            
def get_header_range(sheet, startColumnIndex = 0, endColumnIndex = None):
    if endColumnIndex is None:
        endColumnIndex = sheet.col_count
    return  {
              "sheetId": sheet.id,
              "startColumnIndex": startColumnIndex,
              "endColumnIndex": endColumnIndex,
              "startRowIndex": 0,
              "endRowIndex": 1,
            }           
             
def get_data_range(sheet, startColumnIndex = 0, endColumnIndex = None):
    if endColumnIndex is None:
        endColumnIndex = sheet.col_count
    return  {
              "sheetId": sheet.id,
              "startColumnIndex": startColumnIndex,
              "endColumnIndex": endColumnIndex,
              "startRowIndex": 1,
              "endRowIndex": sheet.row_count,
            }
            
def set_named_ranges_with_errors(sheets, name_ranges, named_range_ids_by_name):
    try: # Attempt to set all named ranges in one request
        exponential_backoff(set_named_ranges, sheets, name_ranges, named_range_ids_by_name, max_n = 2)
    except GoogleAPIException:
        for name, range_ in name_ranges: # Attempt to set named ranges one at a time
            try: 
                exponential_backoff(set_named_ranges, sheets, [(name, range_)], named_range_ids_by_name, max_n = 4)
            except GoogleAPIException:
                print(f"Named Range could not be set.  Please manually set named range '{name}' to Master!{get_column_letter(range_['startColumnIndex'] + 1)}{range_['startRowIndex'] + 1}:{get_column_letter(range_['endColumnIndex'])}{range_['endRowIndex']} in https://docs.google.com/spreadsheets/d/{sheets.id}/edit")
        
            
def set_named_ranges(sheets, name_ranges, named_range_ids_by_name):     
    requests = []  
    for name, range_ in name_ranges: 
        if name in named_range_ids_by_name:
            requests.append({
                  "updateNamedRange": {
                  "fields": "*",
                  "namedRange": {
                      "name": name,
                      "namedRangeId": named_range_ids_by_name[name],
                      "range": range_,
                      }
                  }
            })
        else:
            requests.append({
                  "addNamedRange": {
                  "namedRange": {
                      "name": name,
                      "range": range_,
                      }
                  }
            })
    body = {"requests": requests}
    sheets.batch_update(body)
    
def consent_formatting(sheets, sheet, col_index, required_col):
    requests = []
    ranges = [get_col_data_range(sheet, col_index)]
    requests.append(
    {
      "addConditionalFormatRule": {
        "rule": {
          "ranges": ranges,
          "booleanRule": {
            "condition": {
              "type": "TEXT_STARTS_WITH",
              "values": [
                {
                  "userEnteredValue": "?"
                }
              ]
            },
            "format": {
              "backgroundColor": warning_colour
            }
          }
        },
        "index": 0
      }
    })
    requests.append(
    {
      "addConditionalFormatRule": {
        "rule": {
          "ranges": ranges,
          "booleanRule": {
            "condition": {
              "type": "CUSTOM_FORMULA",
              "values": [
                {
                  "userEnteredValue": "=and(not(isblank(%s2)),isblank(%s2))" % (get_column_letter(required_col), get_column_letter(col_index))
                }
              ]
            },
            "format": {
              "backgroundColor": error_colour
            }
          }
        },
        "index": 1
      }
    })


    body = {"requests": requests}
    return sheets.batch_update(body)
    
def setDropDownFromRange(sheets, sheet, col_index, options_range):
    requests = []
    range_ = get_col_data_range(sheet, col_index)
    requests.append({
      "setDataValidation": {
        "range": range_,
        "rule": {
          "condition": {
            "type": "ONE_OF_RANGE",
            "values": [
              {
                "userEnteredValue": options_range
              }
            ]
          },
          "inputMessage": "My Message",
          "strict": False,
          "showCustomUi": True
        }
      }
    }  
    ) 
    body = {"requests": requests}
    return sheets.batch_update(body)

def batch_colwidths(sheets, sheet, cols):
    requests = []
    for col_index, col_width in cols:
        requests.append({
                  "updateDimensionProperties": {
                    "range": {
                      "sheetId": sheet.id,
                      "dimension": "COLUMNS",
                      "startIndex": col_index - 1,
                      "endIndex": col_index
                    },
                    "properties": {
                      "pixelSize": col_width
                    },
                    "fields": "pixelSize"
                  }
                })

    body = {"requests": requests}
    return sheets.batch_update(body)

def copy_validation_to_col(sheets, source_sheet, source_row, source_col, dest_sheet, dest_col):
    body = {
            "requests": [
                {
                    "copyPaste": {
                        "source": {
                            "sheetId": source_sheet.id,
                            "startRowIndex": source_row,
                            "endRowIndex": source_row + 1,
                            "startColumnIndex": source_col,
                            "endColumnIndex": source_col + 1
                        },
                        "destination": {
                            "sheetId": dest_sheet.id,
                            "startRowIndex": 1,
                            "endRowIndex": dest_sheet.row_count,
                            "startColumnIndex": dest_col -1,
                            "endColumnIndex": dest_col
                        },
                        "pasteType": "PASTE_DATA_VALIDATION"
                    }
                },
            ]
        }
    
    return sheets.batch_update(body)
    
def protectSheet(sheets, sheet, unprotected_cols):
    unprotectedRanges = [{"sheetId": sheet.id,
                          "startRowIndex": 0,
                          "endRowIndex": 1,
                          "startColumnIndex": 0,
                          "endColumnIndex": sheet.col_count,
                          }]
    for col in unprotected_cols:
        unprotectedRanges.append({
                                    "sheetId": sheet.id,
                                    "startRowIndex": 1,
                                    "endRowIndex": sheet.row_count,
                                    "startColumnIndex": col - 1,
                                    "endColumnIndex": col
                                })
    body = {
        "requests": [
            {
                "addProtectedRange": {
                    "protectedRange": {
                        "range": {
                            "sheetId": sheet.id
                        },
                        "unprotectedRanges": unprotectedRanges,
                        "editors": {
                            "domainUsersCanEdit": False,
                            "users": "mjg-python@wirral-green-party-data.iam.gserviceaccount.com"
                        },
                        "warningOnly": False
                    }
                }
            }
        ]
    }
    sheets.batch_update(body)
    
def dtest():
    class r:
        def __init__(self):
            self.text = "Test"
    raise gspread.exceptions.APIError(r())
    
class GoogleAPIException(Exception):
    pass
    
def exponential_backoff(f, *args, max_n = 22, **kwargs):
    try:
        return f(*args, **kwargs)
    except gspread.exceptions.APIError as myerror:
        exponential_backoff_delay(0, 0, f, *args, max_n = max_n, **kwargs)

def exponential_backoff_delay(e, n, f, *args, max_n = 22, **kwargs):
    sleep(2 ** e + random())
    try:
        return f(*args, **kwargs)
    except gspread.exceptions.APIError as myerror:
        print(f"Google API error...({e}) ... {datetime.now()}")
        if e < 6:
            e = e + 1
        if n > max_n:
            raise GoogleAPIException("Google API not accepted a request.")
        exponential_backoff_delay(e, n + 1, f, *args, max_n = max_n, **kwargs)

def write_deleted_record(area, master_headings, new_deleted_records, deleted_date):
    sheet_id, sheet_name = SHEETS[area]
    sheets = get_client().open_by_key(sheet_id)
    sheet = sheets.worksheet("Deleted Electors")   
    deleted_headings = sheet.get('A1:1')[0]
    deleted_date_and_master_headings = ["Deleted Date"] + master_headings
    print (deleted_headings)
    print(deleted_date_and_master_headings)
    for deleted_heading in deleted_headings:
        assert deleted_heading in deleted_date_and_master_headings, 'In Deleted Electors worksheet, "%s" heading does not exist in the Master sheet.  Please apply the changes that were made to Master to the Deleted Electors sheet.' % deleted_heading
    current_deleted_records = sheet.get_all_records()
    sheet.clear()
    sheet.update([deleted_date_and_master_headings], 'A1')
    sheet.update([[deleted_date] + [record[h] for h in master_headings] for record in current_deleted_records] + \
                 [[deleted_date] + [record[h] for h in master_headings] for record in new_deleted_records], 'A2', raw=False)

class UnknownAreaException(Exception):
    pass
    
def read_sheet(area):
    if area not in SHEETS:
        raise UnknownAreaException("Warning area %s not known" % area)
    sheet_id, sheet_name = SHEETS[area]
    sheets = get_client().open_by_key(sheet_id)
    masterSheet = sheets.worksheet(sheet_name)
    headings = masterSheet.get('A1:1')[0]
    formats = []
    for colindex, heading in enumerate(headings, 1):
        column = get_column_letter(colindex)
        if heading in DATE_HEADERS:
            formats.append({"range": f'{column}2:{column}', "format": {"numberFormat": {"type": "DATE", "pattern": "yyyy-mm-dd"}}})   
    exponential_backoff(masterSheet.batch_format, formats)
    
    records = exponential_backoff(masterSheet.get_all_records)
    return sheets, masterSheet, headings, records
    

def add_marks(area, elector_ids, election, postal = []):
    sheets, sheet, headings, records = read_sheet(area)
    if election not in headings:
        start_election_col, end_election_col = find_election_cols(headings)
        
        headings = headings[:end_election_col] + [election] + headings[end_election_col:] #Insert new election column at the end of the elections columns
        for record in records:
            record[election] = record['Elector Number'] in elector_ids
            elector_ids.discard(record['Elector Number'])
            
    else:
        for record in records:
            record[election] = BOOL[record[election]] or (record['Elector Number'] in elector_ids)
            elector_ids.discard(record['Elector Number'])
            
    unfound_worksheet_title = "To Do: %s" % election
    if unfound_worksheet_title not in [wsh.title for wsh in sheets.worksheets()]:
        sheets.add_worksheet(unfound_worksheet_title, rows=1, cols=1)
    for row in sheets.worksheet(unfound_worksheet_title).get_values():
       for cell in row:
           elector_ids.discard(cell)
    unfound_elector_ids = list(elector_ids)
    unfound_elector_ids.sort()
    sheets.worksheet(unfound_worksheet_title).append_rows([[i] for i in unfound_elector_ids])
    
    #Mark postal voters
    for record in records:
        record["PV"] = BOOL[record["PV"]] or record['Elector Number'] in postal
    write_record(area, headings, records)
    
    
    
    
def reorder(area, number_of_records = None):
    sheets, sheet, headings, records = read_sheet(area)
    heading_order_new = ["First Name", "House", "Road", "Round", "Contacted", "LE Rating", "Party", "Mobile", "Consent - Mobile", "Date - Consent - Mobile", "Recorded By - Consent - Mobile", "Consent - WhatsApp", "Date - Consent - WhatsApp", "Recorded By - Consent - WhatsApp", "Landline", "Consent - Landline", "Date - Consent - Landline", "Recorded By - Consent - Landline", "Email", "Consent - Email", "Date - Consent - Email", "Recorded By - Consent - Email", "Poster", "Board", "DNC", "DND", "Notes", "GE Rating", "Chosen Name", "Full Name", "Elector Number"] + [x for x in headings if x[:2] in ["LE", "GE", "MM"] and len(x) <= 6] + ["Address Type", "PV", "Added Date", "Elector Name", "Elector Markers", "Elector DOB", "Previous Elector Number", "Address1", "Address2", "Address3", "Address4", "Address5", "Address6", "Postcode", "AddressFull", "Last Contact", "DNC Date", "DND Date", "Poster Request Date", "Board Request Date"]
    print(heading_order_new)
    print("Missing", set(headings) - set(heading_order_new))
    print("New", set(heading_order_new) - set(headings))
    
    #Set defaults for new columns
    create_data_for_new_cols(records)  
    
    if number_of_records:
        records = records[:number_of_records]    
    write_record(area, heading_order_new, records)
    
def create_data_for_new_cols(records):
    for record in records:
        if "Contacted" not in record:
            record["Contacted"] = ""
        if "DNC Date" not in record:
            if "DNC" in record and record["DNC"]=="TRUE":
                record["DNC Date"] = date.today().strftime("%Y-%m-%d")
            else:
                record["DNC Date"] = ""
        if "DND Date" not in record:
            if "DND" in record and record["DND"]=="TRUE":
                record["DND Date"] = date.today().strftime("%Y-%m-%d")
            else:
                record["DND Date"] = ""
        if "DND" not in record:
            record["DND"] = "FALSE"
        for comm_medium, comm_record, _ in MEDIUMS_RECORDS:
            if f"Consent - {comm_medium}" not in record:
                if record[comm_record] == "":
	                record[f"Consent - {comm_medium}"] = ""
                else:
	                record[f"Consent - {comm_medium}"] = "?Legacy: Consent unknown"
            if f"Date - Consent - {comm_medium}" not in record:
                record[f"Date - Consent - {comm_medium}"] = ""
            if f"Recorded By - Consent - {comm_medium}" not in record:
                record[f"Recorded By - Consent - {comm_medium}"] = "" 
    
def update_rounds(area):
    master_sheets, master_sheet, master_headings, master_records = read_sheet(area)
    addresses = [(record["House"], record["Road"], record["AddressFull"]) for record in master_records]
    addresses = list(OrderedDict.fromkeys(addresses)) # Remove duplicate addresses
    sheet_id, sheet_name = ROADSHEETS[area]
    sheets = get_client().open_by_key(sheet_id)
    oldsheet = sheets.worksheet(sheet_name)
    oldsheet.update_title(f"Old{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    headings = ["House", "Road", "AddressFull"]
    sheet = sheets.add_worksheet(title=sheet_name, rows=len(addresses) + 1, cols=len(headings),index=0)
    sheet.update([headings], 'A1')
    sheet.update([record for record in addresses], 'A2', raw=False) 
    name_ranges = [("Roads", get_data_range(sheet, 1, 2)), 
                   ("HouseRoads", get_data_range(sheet, 0, 2)),
                   ("AddressFull", get_data_range(sheet, 2, 3))]
    current_name_ranges = sheets.list_named_ranges()
    exponential_backoff(set_named_ranges, sheets, name_ranges, dict([(x["name"], x["namedRangeId"]) for x in current_name_ranges]))
    
def write_addresses(area, record_list):
    for n, record in enumerate(record_list):
        make_House_Road_AddressFull(record, area)
    addresses = [(record["House"], record["Road"], record["AddressFull"]) for record in record_list]
    addresses = list(OrderedDict.fromkeys(addresses)) # Remove duplicate addresses
    sheet_id, sheet_name = ROADSHEETS[area]
    sheets = get_client().open_by_key(sheet_id)
    oldsheet = sheets.worksheet(sheet_name)
    oldsheet.update_title(f"Old{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    headings = ["House", "Road", "AddressFull"]
    sheet = sheets.add_worksheet(title=sheet_name, rows=len(addresses) + 1, cols=len(headings),index=0)
    sheet.update([headings], 'A1')
    sheet.update([record for record in addresses], 'A2', raw=False) 
    name_ranges = [("Roads", get_data_range(sheet, 1, 2)), 
                   ("HouseRoads", get_data_range(sheet, 0, 2)),
                   ("AddressFull", get_data_range(sheet, 2, 3))]
    current_name_ranges = sheets.list_named_ranges()
    exponential_backoff(set_named_ranges, sheets, name_ranges, dict([(x["name"], x["namedRangeId"]) for x in current_name_ranges]))
