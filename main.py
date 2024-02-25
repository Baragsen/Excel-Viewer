import tkinter as tk
from tkinter import ttk
import customtkinter 
import excel_reader
import os.path

app = tk.Tk()
app.geometry("800x600")
app.tk.call("source" , "forest-dark.tcl")

style = ttk.Style(app)
style.theme_use("forest-dark")

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


def viewer(frame , path): 
    if os.path.exists(path) and os.path.splitext(path)[1] == ".xlsx":
        frame.destroy()
        excel_viewer_page(path)



def excel_viewer_page(path):
    insert_frame = customtkinter.CTkScrollableFrame(master=app)
    insert_frame.grid(column=0 , row=0 , sticky = "ns")
    app.grid_rowconfigure(0, weight=1) 

    treeframe = ttk.Frame(app)
    treeframe.grid(column= 1  , row=0 , sticky = "nswe")
    app.grid_rowconfigure(0, weight=1) 
    app.grid_columnconfigure(1, weight=1) 

    tree_y_Scrollbar = ttk.Scrollbar(treeframe)
    tree_y_Scrollbar.pack(side="right",fill="y")

    # tree_x_Scrollbar = ttk.Scrollbar(treeframe)
    # tree_x_Scrollbar.pack(side="bottom",fill="x")

    # treeview = ttk.Treeview(treeframe,show='headings' , columns=excel_reader.get_columns(path) , height=)
    entries = [customtkinter.CTkEntry(insert_frame, placeholder_text=column).pack(side="top", pady=10) for column in excel_reader.get_columns(path)]


first_page()
app.mainloop()
