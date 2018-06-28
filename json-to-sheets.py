import argparse
import json
import gspread
from gspread.utils import rowcol_to_a1
from oauth2client.service_account import ServiceAccountCredentials


class GoogleSheetDAL:
    """
    Assumes a worksheet with the first column being designated for date entries. First Row is for item headers. 
    All other rows are data
    """

    def __init__(self, key_file, spreadsheet_id, worksheet_name=None):
        self.google_client = self.get_google_client(key_file)
        self.spreadsheet = self.get_spreadsheet(self.google_client, spreadsheet_id)
        self.data_worksheet = self.get_data_worksheet(self.spreadsheet, worksheet_name)

    @staticmethod
    def get_google_client(key_file):
        scope = ['https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(key_file, scope)
        gc = gspread.authorize(credentials)
        return gc

    @staticmethod
    def get_spreadsheet(google_client, spreadsheet_id):
        spreadsheet = google_client.open_by_key(spreadsheet_id)
        return spreadsheet

    @staticmethod
    def get_data_worksheet(spreadsheet, worksheet_name):
        if worksheet_name is None:
            worksheet = spreadsheet.get_worksheet(0)
        else:
            worksheet = spreadsheet.worksheet(worksheet_name)
        return worksheet

    def find_date_row(self):
        """
        Returns the row number to place the new date timestamp in (first empty row).
        Also performs a consistency check to ensure the first column is for dates.
        """
        date_row = [x for x in self.data_worksheet.col_values(1) if x]
        if "Date" not in date_row[0]:
            raise ValueError("Date header missing from worksheet")
        return len(date_row) + 1

    def get_row_headers(self):
        """
        Returns the first row of the sheet. This is used to find which column items are in.
        Also performs a consistency check to ensure the first column is for dates.
        """
        row_headers = [x for x in self.data_worksheet.row_values(1) if x]
        if "Date" not in row_headers[0]:
            raise ValueError("Date header missing from worksheet")
        return row_headers

    def append_data(self, data):
        """
        Takes an dictionary object and stores it in the spreadsheet
        :param data: Must have a "Date" element. Example: {'Date': '1970-01-01', 'Column1': 'Value1', 'Column2': 'Value2'}
        :return: Nothing
        """
        headers = self.get_row_headers()
        new_row_index = self.find_date_row()
        if self.data_worksheet.row_count < new_row_index:
            self.data_worksheet.add_rows(new_row_index)

        # Check to ensure all the items are in the header, so I can then do a bulk insert. Add them if missing
        new_header = False
        for column_name, value in data.items():
            if column_name not in headers:
                headers.append(column_name)
                new_header = True
        new_column_count = len(headers) - self.data_worksheet.col_count
        if new_column_count > 0:
            self.data_worksheet.add_cols(new_column_count)
        if new_header:
            start_range = rowcol_to_a1(1, 1)
            end_range = rowcol_to_a1(1, len(headers))
            new_row = self.data_worksheet.range(f"{start_range}:{end_range}")
            for i in range(len(new_row)):
                new_row[i].value = headers[i]
            self.data_worksheet.update_cells(new_row)

        # checkout row headers again to be safe
        headers = self.get_row_headers()

        # checkout the list of cells for bulk update
        start_range = rowcol_to_a1(new_row_index, 1)
        end_range = rowcol_to_a1(new_row_index, len(headers))
        new_row = self.data_worksheet.range(f"{start_range}:{end_range}")

        # Update all the new_row items based off having the same array index as header row
        for column_name, value in data.items():
            index = headers.index(column_name)
            new_row[index].value = value

        # Write results to worksheet
        self.data_worksheet.update_cells(new_row)


def main(keyfile, spreadsheet_id, worksheet_name, data_file):
    """
    Parses json file then stores the data in google sheet
    :return: Nothing
    """

    with open(data_file) as fin:
        data = json.load(fin)

    # Remove keys that are only digits and start with zeros
    keys = [x for x in data.keys()]
    for key in keys:
        if key.isdigit():
            new_key = key.lstrip("0")
            data[new_key] = data[key]
            data.pop(key)

    dal = GoogleSheetDAL(keyfile, spreadsheet_id, worksheet_name)

    dal.append_data(data)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Upload json data to google sheets")
    parser.add_argument('--keyfile', type=str, help="keyfile used for google APIs. See readme for more info", required=True)
    parser.add_argument('--spreadsheetid', type=str, help="Spreadsheet Id to upload data. See readme for more info", required=True)
    parser.add_argument('--worksheetname', type=str, help="Worksheet name (optional). Defaults to first worksheet if empty", required=False)
    parser.add_argument('--datafile', type=str, help="JSON serialized file to upload. This must include a 'Date' field", required=True)

    args = parser.parse_args()

    main(args.keyfile, args.spreadsheetid, args.worksheetname, args.datafile)


