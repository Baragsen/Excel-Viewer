import tkinter as tk
from tkinter import ttk
import customtkinter 
import openpyxl  
import excel_reader
import os.path

app = tk.Tk()
app.geometry("800x600")
app.tk.call("source" , "forest-dark.tcl")

style = ttk.Style(app)
style.theme_use("forest-dark")

def viewer(frame , path): 
    if os.path.exists(path) and os.path.splitext(path)[1] == ".xlsx":
        frame.destroy()
        excel_viewer_page(path)

def insert_data(entries , path ,frames_list):
    entries_values = entries
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    sheet.append(entries_values)
    workbook.save(path)

    for frame in frames_list :
        frame.destroy()
    excel_viewer_page(path)

def first_page(): 
    frame = ttk.Frame(master=app)
    frame.pack(fill="both" , expand=True)

    label = ttk.Label(master=frame, text="Excel Viewer" , font=("Roboto" , 30))
    label.pack(padx=10 , pady=70)

    label = ttk.Label(master=frame, text="Excel File Path : " , font=("Roboto" , 12))
    label.pack(padx=10 , pady=40)


    path_entry = customtkinter.CTkEntry(frame, placeholder_text="Enter Excel File Path" , height=35)
    path_entry.pack(padx=160,pady=10 , fill="both")


    next_button= customtkinter.CTkButton(master=frame , text="Next" ,  width=160, height=50 , command= lambda:viewer(frame , path_entry.get()))
    next_button.pack(padx=10,pady=70 )


def excel_viewer_page(path):
    frames_list= []
    insert_frame = customtkinter.CTkScrollableFrame(master=app)
    insert_frame.grid(column=0 , row=0 , sticky = "ns")
    app.grid_rowconfigure(0, weight=1) 

    treeframe = ttk.Frame(app)
    treeframe.grid(column= 1  , row=0 , sticky = "nswe")
    app.grid_rowconfigure(0, weight=1) 
    app.grid_columnconfigure(1, weight=1) 

    frames_list.append(treeframe)

    tree_y_Scrollbar = ttk.Scrollbar(treeframe , orient="vertical")
    tree_y_Scrollbar.pack(side="right",fill="y")

    tree_x_Scrollbar = ttk.Scrollbar(treeframe , orient="horizontal")
    tree_x_Scrollbar.pack(fill="x" , side="bottom")

    treeview = ttk.Treeview(treeframe,show='headings', xscrollcommand=tree_x_Scrollbar.set ,yscrollcommand= tree_y_Scrollbar.set,  columns=excel_reader.get_columns(path) ,height=excel_reader.get_height(path))
    treeview.pack()
    for column in excel_reader.get_columns(path) :
        treeview.heading(column , text=column)

    for row in excel_reader.get_rows(path):
        treeview.insert("" , tk.END , values=row)
    tree_y_Scrollbar.config(command=treeview.yview)
    tree_x_Scrollbar.config(command=treeview.xview)

    frames_list.append(treeview)

    entries = [customtkinter.CTkEntry(insert_frame, placeholder_text=column) for column in excel_reader.get_columns(path)]
    for entry in entries:
        entry.pack(side="top", pady=10)

    insert_button = ttk.Button(insert_frame ,text='Insert Data' , command=lambda:insert_data([entry.get() for entry in entries] , path , frames_list))
    insert_button.pack(pady=10)

first_page()
app.mainloop()
