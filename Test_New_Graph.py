from openpyxl import *
from openpyxl.styles import PatternFill
from datetime import datetime,time

import string
import random as ran

class Graph():
    def __init__(self,Filename,Sheet):
        self.Filename = Filename
        self.Sheet = Sheet
        self.loading = 0
        self.timer = []
        
    def Poc(self):
    #------------------- open excel file --------------------------------
        try:
            self.wb = load_workbook(self.Filename) #use out
        except FileNotFoundError:
            #print("File not found")
            return self.loading #file not found
        else:
            self.loading = 1
    #----------------------------------------------------------------

    #------------------- access sheet or create sheet --------------------------------    
        if "Graph Timing "+ self.Sheet in self.wb.sheetnames:
            self.ws = self.wb["Graph Timing "+self.Sheet]
        else:
            self.ws = self.wb.create_sheet("Graph Timing "+self.Sheet)
    #----------------------------------------------------------------
    
    #------------------- save excel file --------------------------------      
        try:
            self.wb.save(self.Filename)
        except PermissionError:
            #print("Please, Close the Workbook before continuing")
            return self.loading
        else:
            self.loading = 2
            return self.loading
    #----------------------------------------------------------------

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
    
    def create_timeline(self):
        #------------------- generate time --------------------------------
        start_hour = 0
        #stop_hour = 0
        #step_hour = 1
        start_minute = 0
        stop_minute = 31
        #step_minute = 30
        for hour in range(start_hour,24):
            for minute in range(start_minute,stop_minute+1,30):
                formatted_hour = f"{hour:02d}"
                formatted_minute = f"{minute:02d}"
                self.timer.append(f"{formatted_hour}:{formatted_minute}")
        for hour in range(0,1):
            for minute in range(start_minute,stop_minute,30):
                formatted_hour = f"{hour:02d}"
                formatted_minute = f"{minute:02d}"
                self.timer.append(f"{formatted_hour}:{formatted_minute}")
        
        
        
        insert_pos = 1
        round = len(self.timer)
        while True:
            if round == 0:
                break
            else:
                self.ws.cell(row=1, column=insert_pos+1, value=self.timer[insert_pos-1])
                insert_pos += 1
                round-=1
    #----------------------------------------------------------------