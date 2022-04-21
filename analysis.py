from simple_salesforce import Salesforce
import gspread
from datetime import date, datetime, timedelta
from decouple import config
from operator import itemgetter
import time
import sys
from spread_sheet import *

today = datetime.today()
age = 0


#set what kind of case review

def setup():
    global age
    print("Select one of the options below\n")
    print("A) 2 week Check-in\tB)Historical Cases to Today\n")
    print("C) Custom Period\tQ) Quit")

    select = input("Please Enter Your Choice:  ")
    if select.islower() == True:
        select = select.upper()

    match select:
        case "A":
            age = 14
        case "B":
            start_date = datetime.strptime("2021, 11, 1", "%Y, %m, %d")
            age = (today - start_date).days
        case "C":
            date = input("Please Enter Starting date\nFormat DD MM YYYY:  ")
            start_date = datetime.strptime(date, "%d %m %Y")
            age = (today - start_date).days
        case "Q":
            sys.exit()
    get_sheet()
    return age

#Select spreadsheet


def get_cases():
#Gather Cases
    global age
    sf_user = config("sf_user")
    sf_passwd = config("sf_passwd")
    sf_token = config("sf_token")
    sf = Salesforce(password=sf_passwd, username=sf_user, security_token=sf_token)
    print("Gathering Cases")
    query = sf.query_all_iter(
        "SELECT Id, CaseNumber, BS_Days_Since_Last_Activity__c, RMA_Type__c, Origin, Case_Created_By__c, Status, CreatedDate, Subject FROM Case WHERE BS_Age__c <= {} AND RMA_Type__c != 'Part'".format(age)
    )
    temp_list = []

    for i in query:
        if i.get('Status') == "Repair / Replacement Shipped":
            temp_list.append(i)
        if i.get('Status') == "RMA Resolved":
            temp_list.append(i)
        else:
            continue
    case_list = sorted(temp_list, key=itemgetter('CaseNumber'))
    return case_list



def write_cases(case_list):
    #Write to Spreadsheet
    print("Writting to Sheet")
    print("----------------------\n")
    total_writes = 0
    i = 0
    row = 2

    while i < len(case_list):
        print("----------------------")
        print("Writting Case %s" %(case_list[i]['CaseNumber']))
        try:
            date = datetime.strptime(case_list[i]["CreatedDate"], '%Y-%m-%dT%H:%M:%S.%f+0000')
            WorkSheet.sheet.update_cell(row, 1, date.strftime("%m/%d/%Y"))
            WorkSheet.sheet.update_cell(row, 2, case_list[i]['CaseNumber'])
            WorkSheet.sheet.update_cell(row, 3, case_list[i]['RMA_Type__c'])
            WorkSheet.sheet.update_cell(row, 4, case_list[i]['Subject'])
            WorkSheet.sheet.update_cell(row, 5, case_list[i]['Case_Created_By__c'])
            WorkSheet.sheet.update_cell(row, 7, case_list[i]['Origin'])
            WorkSheet.sheet.update_cell(row, 10, url%(case_list[i]["Id"]))
        except:
            print("----------------------\n")
            print("Quota Reached. Waiting...")
            for remaining in range(100, 0, -1):
                sys.stdout.write("\r")
                sys.stdout.write("{:3d} seconds remaining".format(remaining))
                time.sleep(1)
            print("\nResuming Writes\n")
            continue
        else:
            print("Write Complete")
            print("----------------------\n")
            total_writes += 1
            i += 1
            row += 1
    print("All Writes Done")
    print("%i Total Cases Written"%(total_writes))

def main():
    setup()
    case_list = get_cases()
    write_cases(case_list)


if __name__ == "__main__":
    main()

