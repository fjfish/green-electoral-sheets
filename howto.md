# Jobs to run

* Generate ward sheets
* Run annual update once a year
* Import marked when election data
* Import marked postal when postal election

# How to run create wards file

## Running the test ward file

```
python test_create_ward.py
```

Get id's output and input them into `wirral_config.py`

update_headings.py if sheet has no data in

## Process

If spreadsheet for a ward

  Write resetXXX

else

  import annual update for specific ward

end

Now we have a sheet

If you get an internal API error go to the Electoral Data sheet, edit the named range Master_Headers so that it references the Master sheet instead of the Old... sheet

# Annual update of the ward file

Get the ward file from the ward office. Dump it in this directory as a CSV. The file is currently called CB.csv

Test it's going to work

```bash
python import_annual_update.py CB.csv 25-26 01/12/2025 B
#       ^job                   ^file  ^period.          ^ward (blank for all)
#                                            ^ date data was released (British format)   
```
## Testing

1. Create copy of Electoral Data sheet
2. Allow your Google project account to update the master and deleted protected data
3. If they exist in deleted tab remove LE Knock priority and GE Knock Priority columns
4. Add your project ID to `spreadsheet_functions.py` line 551, `protectSheet` function. 
5. Go to round houses tab and allow access

Change `import_annual_update.py` line to use your copy file.


