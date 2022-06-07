from calendar import month
import datetime
from jugaad_data.nse import bhavcopy_save
import gspread
import csv
import os
def make_new_workSheet(month,year):
    sa  = gspread.service_account(filename="to9-352314-fc2393adfa86.json")
    sh = sa.open("1to9")
    sh.add_worksheet(title=(month+year),rows=1000,cols=26)

def check_wks_empty(wks_name):
    # 0 - empty
    # 1 - not empty
    sa  = gspread.service_account(filename="to9-352314-fc2393adfa86.json")
    sh = sa.open("1to9")

    wks = sh.worksheet(wks_name)

    if wks.acell("A1").value:
        return 1
    else:
        return 0
        
def copy_to_gspread(wks_name , bhav_copy_name):
    csv_path  = "./bhav_copy/"+bhav_copy_name
    csv_file = open(csv_path , "r")
    csv_reader = csv.reader(csv_file)
    data = []

    sa  = gspread.service_account(filename="to9-352314-fc2393adfa86.json")
    sh = sa.open("1to9")

    wks = sh.worksheet(wks_name)
    header =  ["SYMBOL" , "OPEN", "HIGHEST HIGH" , "LOWEST LOW" , "CLOSE"]
    wks.append_row(header)
    data  = []
    c=  0
    for line in csv_reader:
        if c == 0:
            c += 1
            continue
        row = [line[0] , float(line[2]),float(line[3]),float(line[4]),float(line[5])]
        data.append(row)
    wks.append_rows(data)


def update_gspread(wks_name , bhav_copy_name):
    csv_path  = "./bhav_copy/"+bhav_copy_name
    csv_file = open(csv_path , "r")
    csv_reader = csv.reader(csv_file)

    sa  = gspread.service_account(filename="to9-352314-fc2393adfa86.json")
    sh = sa.open("1to9")

    wks = sh.worksheet(wks_name)
    csv_data = []
    c = 0
    for line in csv_reader:
        if c == 0 :
            c+=1
            continue
        row = [line[0] , float(line[2]),float(line[3]),float(line[4]),float(line[5])]
        csv_data.append(row)
    gspred_all_data = wks.get_all_values()
    wks.delete_rows(2,wks.row_count)

    new_data = []
    for gspread_row in gspred_all_data[1:]:
        symbol = gspread_row[0]
        open_val = float(gspread_row[1])
        highest_high = float(gspread_row[2])
        lowest_low = float(gspread_row[3])
        last_close = float(gspread_row[4])
        i = 0
        for csv_row in csv_data:
            csv_symbol  = csv_row[0]
            high = csv_row[2]
            low = csv_row[3]
            close = csv_row[4]
            if csv_symbol == symbol:
                if high > highest_high:
                    highest_high = high
                if low < lowest_low:
                    lowest_low = low
                last_close = close
                del csv_data[i]
                break
            i+=1
        new_data.append([symbol,open_val,highest_high,lowest_low,last_close])
    wks.append_rows(new_data)
    wks.append_rows(csv_data)

def download_csv_from_nse(year , month , day):
    # 0 - not downloaded
    # 1 - downloaded
    try:
        bhavcopy_save(datetime.date(year,month,day) , "./bhav_copy")
        return 1
    except:
        return 0

if __name__ == "__main__":
    today_date_time = datetime.datetime.now()
    day = today_date_time.strftime("%d")
    month_shortForm = today_date_time.strftime("%b")
    month_num = today_date_time.strftime("%m")
    year = today_date_time.strftime("%Y")
    wks_name = month_shortForm+year

    print(today_date_time)

    if int(day) == 1:
        make_new_workSheet(month_shortForm,year)

    #terminate the script if date greater than 9
    if int(day) > 9:
        print("Terminated because date > 9")
        exit()

    #download csv for nse
    download_csv_from_nse_flag = download_csv_from_nse(int(year),int(month_num),int(day))
    if download_csv_from_nse_flag == 0 :
        print("terminated because csv from nse not found")
        exit() # terminate script if csv is not downloaded
    
    bhav_copy_name = "cm"+day+month_shortForm+year+"bhav.csv"
    
    #checking if  wks(worksheet) is empty or not  
    check_wks_empty_flag = check_wks_empty(wks_name)

    if check_wks_empty_flag == 0:
        copy_to_gspread(wks_name , bhav_copy_name)
    else:
        update_gspread(wks_name,bhav_copy_name)
    
    bhav_copy_path = "./bhav_copy/"+bhav_copy_name
    os.remove(bhav_copy_path)
    

    


