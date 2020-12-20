from handleRequests import get_app_data_by_id
import json
from handleFiles import initialise, get_active_sheet_read_file, get_active_sheet_write_file, save_files, write_data, current_line_write_file
import sys
from appstoreconnect import Api
from openpyxl import Workbook, load_workbook

first_line = 3
last_line = 7199

if __name__ == '__main__':

    initialise()

    for i in range(first_line, last_line + 1):
        id = get_active_sheet_read_file()["B" + str(i)].value
        print(id)
        json = get_app_data_by_id(id)

        if json is not None:
            if json["resultCount"] != 0:
                write_data(json)

    save_files()
