from openpyxl import *
import string



def InSheet(SheetName):
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
    az = list(string.ascii_uppercase) # 0 = A ---- len-1(25) = Z -- len = 26
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

InSheet("PyDSheet.xlsx")