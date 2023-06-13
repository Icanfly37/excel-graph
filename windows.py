import tkinter as tk
from tkinter import filedialog
import Graph_Generator as GG

def browse_file():
    global file_path
    file_path = filedialog.askopenfilename()
    excel_path.delete(0, tk.END)  # ลบข้อมูลเก่าในช่องข้อความ (ถ้ามี)
    excel_path.insert(tk.END, file_path)  # แทรกตำแหน่งไฟล์ที่เลือกในช่องข้อความ

def pocess():
    #print(file_path)
    #print(sheet_path.get())
    GG.Poc(file_path,str(sheet_path.get()))

# สร้างหน้าต่างหลัก
window = tk.Tk()
window.title("Graph Generator")
window.geometry('600x400')
window.resizable(0,0)
excel_name = tk.Label(window,text="Excel Location : ",font=("bold",18),fg = "black")
excel_name.place(x = 0,y=44)

# สร้างช่องข้อความสำหรับแสดงตำแหน่งไฟล์ที่เลือก
excel_path = tk.Entry(window, width=25,bd = 3,font=("Regular",14),justify="center")
excel_path.place(x = 180,y=46)

# สร้างปุ่มเพื่อเรียกใช้ file dialog
button_browse = tk.Button(window, text="Browse", command=browse_file)
button_browse.place(x = 470,y=46)

#---------------------------------
sheet_name = tk.Label(window,text="Sheet Name : ",font=("bold",18),fg = "black")
sheet_name.place(x = 21,y=90)

# สร้างช่องข้อความสำหรับแสดงตำแหน่งไฟล์ที่เลือก
sheet_path = tk.Entry(window, width=25,bd = 3,font=("Regular",14),justify="center")
sheet_path.place(x = 180,y=92)


# สร้างปุ่มเพื่อออกจากโปรแกรม
button_exit = tk.Button(window, text="Exit" ,font=("bold",15), command=window.quit)
button_exit.place(x = 450,y=300)

# สร้างปุ่มเพื่อสร้าง Graph
button_gen = tk.Button(window, text="Generate!" ,font=("bold",15), command=pocess)
button_gen.place(x = 250,y=300)

# เริ่มการทำงานของ GUI
window.mainloop()