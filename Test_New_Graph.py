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
        self.color = ["00FF0000","0000FF00","000000FF","00FFFF00","00FF00FF","0000FFFF",
              "00FF0000","0000FF00","000000FF","00FFFF00","00FF00FF","0000FFFF",
              "00008000","00808000","00800080","00008080","00C0C0C0","00808080",
              "009999FF","00993366","00FFFFCC","00CCFFFF","00FF8080","000066CC",
              "00CCCCFF","00FF00FF","00FFFF00","0000FFFF","00800080","00008080",
              "000000FF","0000CCFF","00CCFFFF","00CCFFCC","00FFFF99","0099CCFF",
              "00FF99CC","00CC99FF","00FFCC99","003366FF","0033CCCC","0099CC00",
              "00FFCC00","00FF9900","00FF6600","00666699","00969696","00339966",
              "00993300","00993366","00333399"]
        self.head = []
        
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
    
        self.create_timeline()
        self.ws = self.wb[self.Sheet]
        targetcolumn = ["B","D"]
        self.create_header(targetcolumn)
        
    #------------------- save excel file --------------------------------      
        #try:
        #    self.wb.save(self.Filename)
        #except PermissionError:
            #print("Please, Close the Workbook before continuing")
        #    return self.loading
        #else:
        #    self.loading = 2
        #    return self.loading
    #----------------------------------------------------------------
  
    def create_timeline(self):
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
        self.timer.append(self.timer[0])
        insert_pos = 1
        round = len(self.timer)
        while True:
            if round == 0:
                break
            else:
                self.ws.cell(row=2, column=insert_pos+1, value=self.timer[insert_pos-1])
                insert_pos += 1
                round-=1
    
    def create_header(self,cols):
        for col in cols:
            for cell in self.ws[col]:
                self.head.append(cell.value)
                
    
g = Graph("ABB workshop Graph.xlsx","ABB1")
print(g.Poc())