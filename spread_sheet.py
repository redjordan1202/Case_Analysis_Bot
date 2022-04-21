#Functions and variables used by gspread to write info to the spreedsheet

import gspread
import sys


#important column numbers
date_created_c = 1
case_number_c = 2
RMA_type_c = 3
issue_type_c = 4
case_creator_c = 5
origin_c = 7
link_c = 10

#URL for cases
url = "https://d46000000yxg3eag.lightning.force.com/lightning/r/Case/%s/view"

class WorkSheet:
    spreadsheet = None
    sheet = None


gc = gspread.service_account()


def get_sheet():
    sheet_name = input("Enter Title of sheet:  ")
    try:
        WorkSheet.spreadsheet = gc.open(sheet_name)
    except:
        print("Can not open sheet\n Make sure sheet has been shared to the bot")
        sys.exit()
    else:
        sheet_number = input("Enter Worksheet Index:  ")
        WorkSheet.sheet = WorkSheet.spreadsheet.get_worksheet(int(sheet_number))

