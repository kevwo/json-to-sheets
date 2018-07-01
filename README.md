# json-to-sheets

## Example Usage
`python json_to_sheet.py -h`

`python json_to_sheet.py --keyfile "~/myfolder/samplekeyfile-718f3c7be068.json.txt" --spreadsheetid "1G7gWOyq1OPzsxUiF11eEf8CPxe6vzaUR8FbSYmvz-P8" --worksheetname "Data" --datafile "~/data/myfile.json"`


## Requirements

* Python 3.6+
* Google Drive Account
* Perform initial setup steps (see below)

## Initial Setup

Create a google api key and save the file onto your machine. Store the file path for future use. Refer: http://gspread.readthedocs.io/en/latest/oauth2.html
* Make sure to enable Drive API and Sheets API for your project
* If you get a `Rate limit exceeded. User message: \"Sorry, you have exceeded your sharing quota.\"` error, wait a couple minutes and try again.

Create initial spreadsheet with google api account and share owner rights to your google drive account.:

```
import gspread
from oauth2client.service_account import ServiceAccountCredentials
KEY_FILE = "C:/SOMELOCATION/wow-ah-34852d53f2ce.json.txt"
scope = ['https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(KEY_FILE, scope)
gc = gspread.authorize(credentials)
sh = gc.create('A new spreadsheet')
sh.share('otto@example.com', perm_type='user', role='owner')
```

Relevent Docs:
 - http://gspread.readthedocs.io/en/latest/oauth2.html
 - https://github.com/burnash/gspread
 - https://developers.google.com/sheets/api/quickstart/python


## Known issues
* columns that are only digits that also start with zeros, these zeros will be removed. This is due to sheets auto-formating the cell as a number instead of a string.
  *Refer: https://github.com/burnash/gspread/issues/213
