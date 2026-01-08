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

Run annual update once a year
Import marked when election data
Import marked postal when postal election

If you get an internal API error go to the Electoral Data sheet, edit the named range Master_Headers so that it references the Master sheet instead of the Old... sheet

# Annual update of the ward file

Get the ward file from the ward office. Dump it in this directory as a CSV. The file is currently called CB.csv

Test it's going to work

```bash
python import_annual_update.py CB.csv 25-26 01/12/2025 B
#       ^job                   ^file  ^period.          ^ward (blank for all)
#                                            ^ date data was released (British format)   
```

