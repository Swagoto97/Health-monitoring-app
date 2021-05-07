################################################################################
# excel writing
################################################################################
import openpyxl
import time

# print(cell_add)


#################################################################################
# gui based
################################################################################
from tkinter import *
import time
import tkinter.font as font
from tkinter import messagebox
date = time.strftime("%d-%m-%Y", time.localtime())
time = time.strftime("%I:%M %p", time.localtime())


def submit():
    in_date = entry1.get()
    in_time = entry2.get()
    in_body_tem = entry3.get()
    in_oximeter = entry4.get()
    in_bloodpresure = entry5.get()
    in_bloodsuger = entry6.get()
    xl = openpyxl.load_workbook(
        "F:\Working\Programming\Excel_monitor_data\monitoring.xlsx")
    sheet = xl.active
    sh_row = sheet.max_row+1
    sh_degit = ['A', 'B', 'C', 'D', 'E', 'F']
    cell_add = [i+str(sh_row) for i in sh_degit]
    sheet[f"{cell_add[0]}"] = in_date
    sheet[f"{cell_add[1]}"] = in_time
    sheet[f"{cell_add[2]}"] = in_body_tem
    sheet[f"{cell_add[3]}"] = in_oximeter
    sheet[f"{cell_add[4]}"] = in_bloodpresure
    sheet[f"{cell_add[5]}"] = in_bloodsuger
    xl.save("F:\Working\Programming\Excel_monitor_data\monitoring.xlsx")
    entry1.delete(0, END)
    entry2.delete(0, END)
    entry3.delete(0, END)
    entry4.delete(0, END)
    entry5.delete(0, END)
    entry6.delete(0, END)
    entry3.focus_set()
    messagebox.showinfo("showinfo", "Data save Successfully", parent=top)
    entry1.insert(0, f"{date}")
    entry2.insert(0, f"{time}")


top = Tk()
top.title("Data Monitaring...")
top.geometry('550x410+500+200')
myfont = font.Font(family="Times", weight="bold", size=20)
# Code to add widgets will go here...
l1 = Label(top, text=" Welcome to  Data monitaring app..")
l1.configure(font=('Times', 20, 'bold'))
l1.pack()
l2 = Label(top, text="Date:- ")
l2.configure(font=('Times', 14, 'bold'))
l2.place(x=50, y=80)
l3 = Label(top, text="Time:- ")
l3.configure(font=('Times', 14, 'bold'))
l3.place(x=50, y=120)
l4 = Label(top, text="Body temparature:- ")
l4.configure(font=('Times', 14, 'bold'))
l4.place(x=50, y=160)
l5 = Label(top, text="Oximeter:- ")
l5.configure(font=('Times', 14, 'bold'))
l5.place(x=50, y=200)
l6 = Label(top, text="Blood Preasure:- ")
l6.configure(font=('Times', 14, 'bold'))
l6.place(x=50, y=240)
l7 = Label(top, text="Blood Sugar:- ")
l7.configure(font=('Times', 14, 'bold'))
l7.place(x=50, y=280)
entry1 = Entry(top, width="30", font=("Times", 12), relief='sunken', bd=5)
entry1.insert(0, f"{date}")
entry1.place(x=250, y=80)
entry2 = Entry(top, width="30", font=("Times", 12), relief='sunken', bd=5)
entry2.insert(0, f"{time}")
entry2.place(x=250, y=120)
entry3 = Entry(top, width="30", font=("Times", 12), relief='sunken', bd=5)
entry3.insert(0, "")
entry3.place(x=250, y=160)
entry4 = Entry(top, width="30", font=("Times", 12), relief='sunken', bd=5)
entry4.insert(0, "")
entry4.place(x=250, y=200)
entry5 = Entry(top, width="30", font=("Times", 12), relief='sunken', bd=5)
entry5.insert(0, "")
entry5.place(x=250, y=240)
entry6 = Entry(top, width="30", font=("Times", 12), relief='sunken', bd=5)
entry6.insert(0, "")
entry6.place(x=250, y=280)
button1 = Button(top, text='Enter', command=submit)
button1.place(x=200, y=340)
button1.configure(bd=2, relief='raised')
button1['font'] = myfont


top.mainloop()
