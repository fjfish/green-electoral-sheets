from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

import gspread

ROOT_DIR = "1TgjQzoC45zUEMMNDmEV-lBD-aWbVgMep"
TEMPLATE_ELECTORAL_DATA = "18MrtJ4ET7oHemLM0ZRt2LZdZe6NkNVTMhV6UFz9tiV0"
TEMPLATE_ROUNDS = "19t3H3hVU4l93vRImmJS16W7JMtoQ_0fcD52xz8QQ_pA"

def get_client():
    scopes = [
        'https://www.googleapis.com/auth/drive'
    ]
    creds = Credentials.from_service_account_file("wirral-green-party-data-f15572884310.json", scopes = scopes)
    return creds

def get_user_drive():
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    return drive


def create_folder_in_folder(folder_name, parent_folder_id, service):
    
    file_metadata = {
    'name' : folder_name,
    'parents' : [parent_folder_id],
    'mimeType' : 'application/vnd.google-apps.folder'
    }
    file = service.files().create(body=file_metadata,
                                    fields='id').execute()
    
    return file.get('id')
    
def create_ward(ward_name, map_prefix):
    creds = get_client()
    #creds = creds.with_subject("martin.speleo@gmail.com")
    service = build("drive", "v3", credentials=creds)
    ward_folder_id = create_folder_in_folder(ward_name, ROOT_DIR, service)
    leaflet_round_folder_id = create_folder_in_folder("Leaflet Rounds", ward_folder_id, service)
    marked_register_folder_id = create_folder_in_folder("Marked Registers", ward_folder_id, service)
    door_knock_print_out_id = create_folder_in_folder("Print Outs", ward_folder_id, service)
    maps_id = create_folder_in_folder("Maps", leaflet_round_folder_id, service)
    leaflets_print_out_id = create_folder_in_folder("Print Outs", leaflet_round_folder_id, service)
    drive = get_user_drive()
    files = drive.auth.service.files()
    ed_file = files.copy(fileId=TEMPLATE_ELECTORAL_DATA, body={"parents": [{"id": ward_folder_id}], 'title': f'{ward_name} Electoral Data'} ).execute()
    lr_file = files.copy(fileId=TEMPLATE_ROUNDS, body={"parents": [{"id": leaflet_round_folder_id}], 'title': f'{ward_name} Leaflet Rounds'} ).execute()
    s = gspread.authorize(creds)
    ed_sheets = s.open_by_key(ed_file["id"])
    ed_sheet = ed_sheets.worksheet("Update")
    ed_sheet.update([[door_knock_print_out_id]], 'PrintOutFolder')
    ed_sheet.update([[maps_id]], 'ImagesFolder')
    ed_sheet.update([[map_prefix]], 'ImagePrefix')
    ed_sheet.update([[lr_file["id"]]], 'RoundsSheet')
    lr_sheets = s.open_by_key(lr_file["id"])
    lr_sheet = ed_sheets.worksheet("Update")
    lr_sheet.update([[leaflets_print_out_id]], 'PrintOutFolder')
    lr_sheet.update([[maps_id]], 'ImagesFolder')
    lr_sheet.update([[map_prefix]], 'ImagePrefix')
    
    print("SHEETS")
    print (ed_file["id"])
    print ()    
    print ("ROADSHEETS")
    print (lr_file["id"])
    print ()

create_ward("HoylakeAndMeols", "HoylakeAndMeols_")
create_ward("WestKirbyAndThurstaston", "WestKirbyAndThurstaston_")
