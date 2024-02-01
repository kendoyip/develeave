import cx_Oracle
import sys
import pandas as pd
import datetime

from ..eleave import views as eleave
from flask import Blueprint
from pathlib import Path
from my_app import db,  client
from datetime import date

maintenance = Blueprint('maintenance', __name__)
DEBUG = False

# Global Setting
# Oracle connection
cx_Oracle.init_oracle_client(lib_dir='C:/instantclient_21_6')
# Get python file from instantclient (acc and pw)
sys.path.insert(0, "F:/mmgapp/tool/instantclient")
import config
# Connection
dns_tns = cx_Oracle.makedsn(config.DBname,config.Port, sid=config.SID)
connectionODB = cx_Oracle.connect(config.UserName, config.UserPw, dns_tns)
currentODB = connectionODB.cursor()

# Function
def downloadOriginalData(cursor):

    # Download leave record from ae_eleave_dtl Oracle
    
    cursor.execute("ALTER SESSION SET NLS_DATE_FORMAT = 'MM-DD-YYYY'")

    executeSQLcode = "select * from ae_eleave_dtl where leave_entitle_period LIKE '%2024%'"

    cursor.execute(executeSQLcode)

    result = cursor.fetchall()

    columnNames = [column[0] for column in currentODB.description]

    df = []

    for record in result:
      df.append(dict(zip(columnNames , record)))

    try:
        df = pd.DataFrame(df).to_excel(str(Path(__file__).parent.absolute()) + "/" + "source(apply).xlsx",index=False)
    except:
        pass


    return str(Path(__file__).parent.absolute()) + "/" + "source(apply).xlsx"


def downloadOriginalAppHist(cursor):

    # Download leave record from ae_eleave_app_hist Oracle

    executeSQLcode = ("""
                        with submit_date as
                        (
                        select min(created_date)-1 as min_date from ae_eleave_dtl 
                        where leave_entitle_period LIKE '%2024%'
                        )
                        select ae_eleave_approval_hist.* from ae_eleave_approval_hist, submit_date
                        where approval_date >= submit_date.min_date
                        and request_type = 'Apply' and approval_status = 'Approved'
                     """)
    
    cursor.execute(executeSQLcode)

    result = cursor.fetchall()

    columnNames = [column[0] for column in currentODB.description]

    df = []

    for record in result:
      df.append(dict(zip(columnNames , record)))

    try:
        df = pd.DataFrame(df).to_excel(str(Path(__file__).parent.absolute()) + "/" + "source(approve).xlsx",index=False)
    except:
        pass

    return str(Path(__file__).parent.absolute()) + "/" + "source(approve).xlsx"      


def apitransition(input, approvalinput):

    ##apply format :{"year" : 2022, "racf": "NF1KWY", "type": "ANNUAL LEAVE",
    ##               "applying": [   {"startDate": "2022-07-19", "startTime": "AM", "endDate": "2022-07-19", "endTime": "PM"},
    ##                               {"startDate" : "2022-07-20", "startTime": "PM", "endDate": "2022-07-20", "endTime": "PM"},
    ##                               { "startDate" : "2022-08-01", "startTime": "AM", "endDate" : "2022-08-02", "endTime": "AM"}
    ##                           ],
    ##               "applyingScreen": [   {"startDate": "2022-07-19", "startTime": "Full Day", "endDate": "2022-07-19", "endTime": "Full Day"},
    ##                                     {"startDate": "2022-07-20", "startTime": "Half Day - PM", "endDate": "2022-07-20", "endTime": "Half Day - PM"},
    ##                                     {"startDate": "2022-08-01", "startTime": "Full Day", "endDate": "2022-08-02", "endTime": "Half Day - AM"}
    ##                                  ],
    ##               "superUser": false }

    df = pd.read_excel(input, usecols="A,B,D,E,G,I,J,K,L,M,N,O,T")

    df.sort_values(by="APPLICANT_RACF", ascending=False, inplace=True)

    done_list = []

    ## Transform to new format
    ##applyslot = {"year", "racf", "type", "applying", "applyingScreen", "superUser"}
    apply = [ ]
    for n in range(len(df)):

        ref_no = str(df["REF_NO"].values[n])


        if str(df["REQUEST_TYPE"].values[n]) == "Apply" and (ref_no not in done_list):
            
            # Make leave category to code:
            leavecode = ""
            if df["LEAVE_CATEGORY"].values[n] == "Annual Leave": leavecode = "LVE01"
            elif df["LEAVE_CATEGORY"].values[n] == "Casual / Festivities Leave": leavecode = "LVE02"
            elif df["LEAVE_CATEGORY"].values[n] == "Work From Home": leavecode = "LVE03"
            elif df["LEAVE_CATEGORY"].values[n] == "Sick Leave - With Medical Cert.": leavecode = "LVE04"
            elif df["LEAVE_CATEGORY"].values[n] == "Sick Leave - No Medical Cert.": leavecode = "LVE05"

            # Make AM / PM / Full Day
            if "AM" in df["LEAVE_START_AMPM"].values[n]: start_am_pm = "AM"
            elif "PM" in df["LEAVE_START_AMPM"].values[n]: start_am_pm = "PM"
            else: start_am_pm = start_am_pm = "AM"

            if "AM" in df["LEAVE_END_AMPM"].values[n]: end_am_pm = "AM"
            elif "PM" in df["LEAVE_END_AMPM"].values[n]: end_am_pm = "PM"
            else: end_am_pm = end_am_pm = "PM"     

            # Make applying array
            applying = [ ]

            for s in range(len(df)):

                if df["REF_NO"].values[n] == df["REF_NO"].values[s]:
                    
                    applying_slot = {            
                                    "startDate": (pd.to_datetime(df["LEAVE_START"].values[s]).date()).strftime('%Y-%m-%d'),
                                    "startTime": start_am_pm,
                                    "endDate": (pd.to_datetime(df["LEAVE_END"].values[s]).date()).strftime('%Y-%m-%d'),
                                    "endTime": end_am_pm
                                    }
                    applying.append(applying_slot)

            applyingScreen = [ ]

            for s in range(len(df)):

                if df["REF_NO"].values[n] == df["REF_NO"].values[s]:

                    applyingScreen_slot = {
                                          "startDate": (pd.to_datetime(df["LEAVE_START"].values[s]).date()).strftime('%Y-%m-%d'),
                                          "startTime": start_am_pm,
                                          "endDate": (pd.to_datetime(df["LEAVE_END"].values[s]).date()).strftime('%Y-%m-%d'),
                                          "endTime": end_am_pm
                                          }

                    applyingScreen.append(applyingScreen_slot)

            approval_df = pd.read_excel(approvalinput, usecols="A,C,D,G")

            approver1 = ""
            approval_date1 = ""
            approver2 = ""
            approval_date2 = ""            
            approver3 = ""
            approval_date3 = ""     

            for m in range(len(approval_df)):
                if str(approval_df["REF_NO"].values[m]) == str(df["REF_NO"].values[n]):
                    if str(approval_df["APPROVAL_SEQ"].values[m]) == "1":
                        approver1 = approval_df["APPROVER_RACF"].values[m]
                        approval_date1 = approval_df["APPROVAL_DATE"].values[m]
                        approval_date1 = (pd.to_datetime(approval_date1).date()).strftime('%Y-%m-%d')
                    elif str(approval_df["APPROVAL_SEQ"].values[m]) == "2":
                        approver2 = approval_df["APPROVER_RACF"].values[m]
                        approval_date2 = approval_df["APPROVAL_DATE"].values[m]
                        approval_date2 = (pd.to_datetime(approval_date2).date()).strftime('%Y-%m-%d')
                    elif str(approval_df["APPROVAL_SEQ"].values[m]) == "3":
                        approver3 = approval_df["APPROVER_RACF"].values[m]
                        approval_date3 = approval_df["APPROVAL_DATE"].values[m]
                        approval_date3 = (pd.to_datetime(approval_date3).date()).strftime('%Y-%m-%d')

            apply_slot = {
                         "year": 2023, 
                         "racf": df["APPLICANT_RACF"].values[n], 
                         "type": leavecode, 
                         "applying": applying, 
                         "applyingScreen": applyingScreen, 
                         "superUser": False,
                         "status": df["STATUS"].values[n],
                         "submitdate": (pd.to_datetime(df["CREATED_DATE"].values[n]).date()).strftime('%Y-%m-%d'),
                         "approver1": approver1,
                         "approval_date1": approval_date1,
                         "approver2": approver2,
                         "approval_date2": approval_date2,
                         "approver3": approver3,
                         "approval_date3": approval_date3
                        }
            apply.append(apply_slot)

            done_list.append(df["REF_NO"].values[n])
    
    return apply

# Main part
print (" *** Starting transition for eleave record .... ***")

# Download leave record from Oracle first
data = downloadOriginalData(currentODB)

# Download approval record from Oracle
data_approval = downloadOriginalAppHist(currentODB)

# Transform the old data format to input api format to apply leave
input = apitransition(data, data_approval)

#print (input) ## testing
#sys.exit() ## testing

## Apply input leave
## Call eleave function : applyleave 
##result = views.listLeave(input)
df = pd.DataFrame(input)
## Create new column result, new ref number
df["ref_no"] = ""
df["result"] = ""
index = 1


for x in range(len(input)):

    print ("Applying ... " + str(input[x]))
    print ("------------------------------" + str(x) + "/" + str((len(df))) + "-----------------------------------------------")

    result = eleave.applyLeave(input[x])
    #result = {"pass": True, "error_message" : None, "result": [], "Status_code": 200}

    if x > 0 and str(input[x]["racf"]) == str(input[x-1]["racf"]):
        index = index + 1
    else:
        index = 1

    df["ref_no"].values[x] = str(input[x]["year"]) + "00" + str(index)

    if result["Status_code"] == 200:
        df["result"].values[x] = "Passed"
    else:
        df["result"].values[x] = "Failed : " + str(result["error_message"])
    
    print (result)
    print ("-----------------------------------------------------------------------------")

# Create log file ( + new format reference# and input status)

df.to_excel(str(Path(__file__).parent.absolute()) + "/" + "log(apply).xlsx", index=False)

# Approve input leave
#Global Constant
status = list(db["eleave_maintenance"].find({"table": { "$eq" : "globalConstant"}}))
sdf = pd.DataFrame(status)
eleaveDtl = db["eleave_dtl"]

# input = {"refNo": "REG2023001BHC", "racf": approver - RACF, "action": sdf['gcActionApprove'][0]}
# result = views.changeStatus(input)

df2 = pd.DataFrame(input)
df2["ref_no"] = ""
df2["result"] = ""

for x in range(len(df)):

    staffRecord = list(eleaveDtl.find ( {"staff.racf" : { '$regex' : df["racf"].values[x], '$options' : "i"} } ) )

    new_ref_no = str(staffRecord[0]["staff"]["hr_office"]) + str(df["ref_no"].values[x]) + str(df["racf"].values[x])[3:]

    approver_list = [staffRecord[0]["staff"]["approver1"], staffRecord[0]["staff"]["approver2"], staffRecord[0]["staff"]["approver3"]]

    localTime = date.today().strftime("%a %b %d %Y")

    for appno, approver in enumerate(approver_list):
        
        #if df["status"].values[x] == "Approved" and len(approver) > 1 and df["result"].values[x] == "Passed":
        if len(approver) > 1 and df["result"].values[x] == "Passed" and df[f"approval_date{appno+1}"].values[x] != "":
        
            approveInput = {"year": df["year"].values[x], "refNo": new_ref_no, "racf": approver, "localTime": localTime  ,"action": sdf['gcActionApprove'][0]}

            result = eleave.changeStatus(approveInput)
            #result = {"pass": True, "error_message" : None, "result": [], "Status_code": 200}

            df2["ref_no"].values[x] = new_ref_no

            if result["Status_code"] == 200:
                df2["result"].values[x] = "Passed"
            else:
                df2["result"].values[x] = "Failed in approval : " + approver + ", result : " + str(result["error_message"])

            print ("Approving ... " + str(approveInput) )
            print ("------------------------------" + str(x) + "/" + str((len(df))) + "-----------------------------------------------")
            print (result)
        
        else:

            continue

    try:

        #if result["Status_code"] == 200 and df["status"].values[x] == "Approved":
        if result["Status_code"] == 200:

            # Update Submit Date and Approval Date if it is successful application
                
                changeDateInput = {"ref_no": str(df["ref_no"].values[x]), "racf": str(df["racf"].values[x]), "submitdate": df["submitdate"].values[x] , "approver1": df["approver1"].values[x] , "approval_date1": df["approval_date1"].values[x] , "approver2": df["approver2"].values[x] , "approval_date2": df["approval_date2"].values[x] , "approver3": df["approver3"].values[x],  "approval_date3": df["approval_date3"].values[x]}
                
                print ("------------------------------------------------------------------------")
                print ("Changing ... " + str(changeDateInput))
                
                result = eleave.apiChangeLeaveRecordDate(changeDateInput)
                #result = {"pass": True, "error_message" : None, "result": [], "status_code": 200}
                print (result)

                if result["status_code"] == 200:
                    df2["result"].values[x] = "Passed"
                else:
                    df2["result"].values[x] = "Failed in changing date : " + approver + ", result : " + str(result["error_message"])

    except Exception as error:

        df2["result"].values[x] = "Failed in changing date : " + error

        print ("Failed in changing date : " + error)
    

df2.to_excel(str(Path(__file__).parent.absolute()) + "/" + "log(approve).xlsx", index=False)

