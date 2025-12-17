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


def get_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets"
    ]
    creds = Credentials.from_service_account_file("wirral-green-party-data-f15572884310.json", scopes = scopes)
    return gspread.authorize(creds)

sheets = get_client().open_by_key("1_HMUM6fbhD6P5696wbBpaAcIaLSB2jROfcBAddedjB8")

sheet = sheets.worksheet("Master")
print("sheet.id", sheet.id)

current_name_ranges = sheets.list_named_ranges()
lookup = dict([(x["name"], x["namedRangeId"]) for x in current_name_ranges])
print(lookup["Master_Headers"])

#requests = [{'updateNamedRange': {'fields': '*', 'namedRange': {'name': 'Master_Headers', 'namedRangeId': '1202757141', 'range': {'sheetId': 710290393, 'startColumnIndex': 1, 'endColumnIndex': 5, 'startRowIndex': 1, 'endRowIndex': 2}}}}]

name = "MyNamedRange"

requests = [{'updateNamedRange': {'fields': '*', 'namedRange': {'name': name, 'namedRangeId': lookup[name], 'range': {'sheetId': sheet.id, 'startColumnIndex': 1, 'endColumnIndex': 5, 'startRowIndex': 1, 'endRowIndex': 2}}}}]


body = {"requests": requests}

try:
    sheets.batch_update(body)
except gspread.exceptions.APIError as MyError:
    print(MyError)
    
print("WORKED :)")

name = "Master_Headers"

requests = [{'updateNamedRange': {'fields': '*', 'namedRange': {'name': name, 'namedRangeId': lookup[name], 'range': {'sheetId': sheet.id, 'startColumnIndex': 1, 'endColumnIndex': 5, 'startRowIndex': 1, 'endRowIndex': 2}}}}]


body = {"requests": requests}

try:
    sheets.batch_update(body)
except gspread.exceptions.APIError as MyError:
    print(MyError)

