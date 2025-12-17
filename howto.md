# How to run this

These are very rough notes

# Create project in google

* Add the googleapiclient to the project.
* Enable Google Sheets API

# Setting up the venv

## 1. Create it

```
python3 -m venv ../electoral-venv
```

## 2. Populate it

Follow the instructions

```
pip install google_auth_oauthlib
pip install google-api-python-client
pip install pydrive
pip install gspread
```

# App credentials

Can't remember how we got this file, but it's somewhere in the app settings. 

# Client secrets file

This is copied from a help file and the steps aren't quite right, but you get the general idea. If I recall correctly I previously had to turn on OAuth for the project in one of the menus.

The client_secrets.json file is typically required for authenticating with Google APIs (such as Google Sheets, Drive, etc.). You can obtain it by following these steps:

1. Go to the Google Cloud Console: https://console.cloud.google.com/
2. Create a new project (or select an existing one).
3. Enable the relevant API (e.g., Google Sheets API, Google Drive API).
4. Go to "APIs & Services" > "Credentials".
5. Click "Create Credentials" > "OAuth client ID".
6. Configure the consent screen if prompted.
7. Choose "Desktop app" or "Other" as the application type.
8. Download the generated client_secrets.json file.
9. Place the file in your project directory as required by your code.

# Running the test ward file

```
python3 test_create_ward.py
```

Get id's output and input them into `wirral_config.py`

update_headings.py if sheet has no data in

# Process

If spreadsheet for a ward

  Write resetXXX

else

  import annual update for specific ward

end

Now we have a sheet

Run annual update once a year
Import marked when election data
Import marked postal when postal election

If you get an internal API error go to the Electoral Data sheet, edit the named range Master_Headers so that it references the Master sheet instead of the Old... sheet