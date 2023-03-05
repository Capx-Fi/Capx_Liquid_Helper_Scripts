import pandas as pd
import csv
import datetime
from datetime import timezone
import re

excel_sheet = "Sheet.xlsx"
excel_data = pd.read_excel(excel_sheet)

data = pd.DataFrame(excel_data, columns=["Wallet Address","Amount","Start Date for Vesting","End Date for Vesting"])

def getMonth(month) :
    if (month == "jan" or month == "january") : 
        return 1
    elif (month == "feb" or month == "february"):
        return 2
    elif (month == "mar" or month == "march"):
        return 3
    elif (month == "apr" or month == "april"):
        return 4
    elif (month == "may"):
        return 5
    elif (month == "jun" or month == "june"):
        return 6
    elif (month == "jul" or month == "july"):
        return 7
    elif (month == "aug" or month == "august"):
        return 8
    elif (month == "sep" or month == "september"):
        return 9
    elif (month == "oct" or month == "october"):
        return 10
    elif (month == "nov" or month == "november"):
        return 11
    elif (month == "dec" or month == "december"):
        return 12
    else:
        return 0

def getTimestamp(date) :
    string = date.split()
    month = getMonth(string[1].lower())
    if (month > 0):
        day = string[0]
        year = datetime.date.today().year
        dt = datetime.datetime(int(year), int(month), int(day))
        timestamp = dt.replace(tzinfo=timezone.utc).timestamp()
        return int(timestamp)
        

def isValidAddress(address):
    regex = r"^0x[a-fA-F0-9]{40}$"
    if re.match(regex, address, re.IGNORECASE) :
        return True
    else:
        return False

def getDate(timestamp):
    dt_object = datetime.datetime.fromtimestamp(timestamp).strftime('%d-%m-%Y')
    return dt_object

new_data = []
current_timestamp = (int((datetime.datetime.now().timestamp()) // 86400)+1) * 86400
sr_no = 1
for address,amount,start_date,end_date in zip(data["Wallet Address"],data["Amount"],data["Start Date for Vesting"],data["End Date for Vesting"]):
    if (isValidAddress(address)) :
            start_time = getTimestamp(start_date)
            end_time = getTimestamp(end_date)
            total_difference = ((end_time - start_time) // 86400) + 1
            actual_difference = ((end_time - current_timestamp) // 86400) + 1
            if (actual_difference > 0) :
                each_day_amount = amount / total_difference
                amount_provided = 0
                for i in range(0, actual_difference):
                    if i+1 == actual_difference:
                        new_data.append([sr_no, address, getDate(current_timestamp+(i*86400)), amount-amount_provided])
                        sr_no+=1
                    elif i == 0:
                        new_data.append([sr_no, address, getDate(current_timestamp+(i*86400)), each_day_amount * ((total_difference - actual_difference) + 1)])
                        amount_provided+=(each_day_amount * ((total_difference - actual_difference) + 1))
                        sr_no+=1
                    else:
                        new_data.append([sr_no, address, getDate(current_timestamp+(i*86400)), each_day_amount])
                        amount_provided+=each_day_amount
                        sr_no+=1
            else :
                new_data.append([sr_no, address,getDate(current_timestamp), amount])
                sr_no+=1

fields = "Sr. No.,Address,Date(DD-MM-YYYY),Amount of Tokens".split(",")
filename = "capx_sheet.csv"

with open(filename, 'w') as csvfile:
    # creating a csv writer object 
    csvwriter = csv.writer(csvfile) 
        
    # writing the fields 
    csvwriter.writerow(fields) 
        
    # writing the data rows 
    csvwriter.writerows(new_data)

df_new = pd.read_csv('capx_sheet.csv')
 
# saving xlsx file
GFG = pd.ExcelWriter('capx_sheet.xlsx')
df_new.to_excel(GFG, index=False)
 
GFG.save()