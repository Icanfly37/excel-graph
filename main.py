from openpyxl import *
import string


def gentime(l,start_hour,stop_hour,step_hour,start_minute,stop_minute,step_minute):
    for hour in range(start_hour,stop_hour,step_hour):
        for minute in range(start_minute,stop_minute,step_minute):
            # Format the hour and minute with leading zeros
            formatted_hour = f"{hour:02d}"
            formatted_minute = f"{minute:02d}"
            
            # Print or use the formatted hour and minute
            l.append(f"{formatted_hour}:{formatted_minute}")
            #print(f"{formatted_hour}:{formatted_minute}")
    return l

# start_hour 0,stop_hour 23(+1),step_hour 1,start_minute 0,stop_minute 30(+1),step_minute 30

def time_cal(start_hour,end_hour,start_min,end_min):
    l = []
    if end_hour == 0:
        if start_hour == 0 or start_hour > 0:
            l = gentime(l,start_hour,24,1,start_min,end_min+1,30)
            l = gentime(l,0,1,1,start_min,end_min,30)
    elif end_hour - start_hour > 0:
        l = gentime(l,start_hour,end_hour+1,1,start_min,end_min+1,30)
    else:
        l = gentime(l,start_hour,24,1,start_min,end_min+1,30)
        l = gentime(l,0,end_hour+1,1,start_min,end_min+1,30)
    return l


def main(SheetName):
#------------------- open excel file --------------------------------
    try:
        wb = load_workbook(SheetName)
    except FileNotFoundError:
        print("Sheet not found")
#----------------------------------------------------------------

#------------------- access sheet or create sheet --------------------------------    
    if "graph" in wb.sheetnames:
        ws = wb["graph"]
    else:
        ws = wb.create_sheet("graph")
#----------------------------------------------------------------

#------------------- generate A-Z matrix --------------------------------
    #az = list(string.ascii_uppercase) # 0 = A ---- len-1(25) = Z -- len = 26
    #col_time = 0
#----------------------------------------------------------------

#------------------- generate time --------------------------------
    start_hour = 0
    end_hour = 0
    #step_hour = 1
    start_min = 0
    end_min = 31
    #step_minute = 30

    timer = time_cal(start_hour,end_hour,start_min,end_min)
    insert_pos = 1
    round = len(timer)
    while True:
        if round == 0:
            break
        else:
            ws.cell(row=1, column=insert_pos+1, value=timer[insert_pos-1])
            insert_pos += 1
            round-=1
#----------------------------------------------------------------   

#------------------- Do something in sheet --------------------------------
    ##ws["B1"]="Hello Python"
    #if ws[az[1]+"1"] != "0:00" and ws[az[25]+"25"] != "0:00":
    #    time = "0:00"
    #    for i in range(1,26):
    #        for j in range:
    #            ws[az[i]+str(i)] = "0.00"
    #else:
        
        
        
#----------------------------------------------------------------
      
#------------------- save excel file --------------------------------      
    try:
        wb.save(SheetName)
    except PermissionError:
        print("Please, Close the Workbook before continuing")
#----------------------------------------------------------------

main("PyDSheet.xlsx")