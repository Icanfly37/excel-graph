import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox as mb
import Graph_Generator as GG

def browse_file():
    global file_path
    file_path = filedialog.askopenfilename()
    excel_path.delete(0, tk.END)  # ลบข้อมูลเก่าในช่องข้อความ (ถ้ามี)
    excel_path.insert(tk.END, file_path)  # แทรกตำแหน่งไฟล์ที่เลือกในช่องข้อความ

def pocess():
    #print(file_path)
    #print(sheet_path.get())
    try:
        log = GG.Poc(file_path,str(sheet_path.get()))
    except NameError:
        mb.showerror("Excel File","Please, Enter a File, Not Empty!")
    else:
        if log == 0:
            mb.showerror("Excel File","File Not Found")
        elif log == 1:
            mb.showwarning("Save Problem","Please, Close the Workbook before continuing")
        else:
            mb.showinfo("Process","Generate Graph Successful")

# สร้างหน้าต่างหลัก
window = tk.Tk()
window.title("Graph Generator")
window.geometry('600x400')
window.resizable(0,0)
excel_name = tk.Label(window,text="Excel Location : ",font=("bold",18),fg = "black")
excel_name.place(x = 30,y=64)

# สร้างช่องข้อความสำหรับแสดงตำแหน่งไฟล์ที่เลือก
excel_path = tk.Entry(window, width=25,bd = 3,font=("Regular",14),justify="center")
excel_path.place(x = 210,y=66)

# สร้างปุ่มเพื่อเรียกใช้ file dialog
button_browse = tk.Button(window, text="Browse", command=browse_file)
button_browse.place(x = 500,y=66)

#---------------------------------
sheet_name = tk.Label(window,text="Sheet Name : ",font=("bold",18),fg = "black")
sheet_name.place(x = 51,y=110)

# สร้างช่องข้อความสำหรับแสดงตำแหน่งไฟล์ที่เลือก
sheet_path = tk.Entry(window, width=25,bd = 3,font=("Regular",14),justify="center")
sheet_path.place(x = 210,y=112)


# สร้างปุ่มเพื่อออกจากโปรแกรม
button_exit = tk.Button(window, text="Exit" ,font=("bold",15), command=window.quit)
button_exit.place(x = 450,y=300)

# สร้างปุ่มเพื่อสร้าง Graph
button_gen = tk.Button(window, text="Generate!" ,font=("bold",15), command=pocess)
button_gen.place(x = 250,y=300)

# เริ่มการทำงานของ GUI
window.mainloop()