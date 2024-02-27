import tkinter as tk
from tkinter import ttk
import customtkinter 
import openpyxl  
import excel_reader
import os.path

# Create the main application window
app = tk.Tk()
app.geometry("800x600")  # Set window size
app.tk.call("source" , "../forest-dark.tcl")  # Load custom theme

# Set the style and theme for ttk widgets
style = ttk.Style(app)
style.theme_use("forest-dark")

# Function to display the Excel viewer page
def viewer(frame , path): 
    # Check if the file exists and is an Excel file
    if os.path.exists(path) and os.path.splitext(path)[1] == ".xlsx":
        frame.destroy()  # Destroy the current frame
        excel_viewer_page(path)  # Call the excel_viewer_page function with the path

# Function to insert data into the Excel file
def insert_data(entries , path ,frames_list):
    # Get values from the entry fields
    entries_values = entries
    # Open the workbook
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    # Append the values to the sheet
    sheet.append(entries_values)
    # Save the workbook
    workbook.save(path)

    # Destroy all frames in the list
    for frame in frames_list :
        frame.destroy()
    # Display the excel_viewer_page with the updated data
    excel_viewer_page(path)

# Function to display the first page
def first_page(): 
    # Create a frame
    frame = ttk.Frame(master=app)
    frame.pack(fill="both" , expand=True)

    # Add labels and entry fields
    label = ttk.Label(master=frame, text="Excel Viewer" , font=("Roboto" , 30))
    label.pack(padx=10 , pady=70)

    label = ttk.Label(master=frame, text="Excel File Path : " , font=("Roboto" , 12))
    label.pack(padx=10 , pady=40)

    path_entry = customtkinter.CTkEntry(frame, placeholder_text="Enter Excel File Path" , height=35)
    path_entry.pack(padx=160,pady=10 , fill="both")

    # Add a button to proceed
    next_button= customtkinter.CTkButton(master=frame , text="Next" ,  width=160, height=50 , command= lambda:viewer(frame , path_entry.get()))
    next_button.pack(padx=10,pady=70 )

# Function to display the Excel viewer page
def excel_viewer_page(path):
    frames_list= []
    # Create a scrollable frame for insertion
    insert_frame = customtkinter.CTkScrollableFrame(master=app)
    insert_frame.grid(column=0 , row=0 , sticky = "ns")
    app.grid_rowconfigure(0, weight=1) 

    # Create a frame for the treeview
    treeframe = ttk.Frame(app)
    treeframe.grid(column= 1  , row=0 , sticky = "nswe")
    app.grid_rowconfigure(0, weight=1) 
    app.grid_columnconfigure(1, weight=1) 

    frames_list.append(treeframe)

    # Add vertical scrollbar
    tree_y_Scrollbar = ttk.Scrollbar(treeframe , orient="vertical")
    tree_y_Scrollbar.pack(side="right",fill="y")

    # Add horizontal scrollbar
    tree_x_Scrollbar = ttk.Scrollbar(treeframe , orient="horizontal")
    tree_x_Scrollbar.pack(fill="x" , side="bottom")

    # Create a Treeview widget
    treeview = ttk.Treeview(treeframe,show='headings', xscrollcommand=tree_x_Scrollbar.set ,yscrollcommand= tree_y_Scrollbar.set,  columns=excel_reader.get_columns(path) ,height=excel_reader.get_height(path))
    treeview.pack()
    for column in excel_reader.get_columns(path) :
        treeview.heading(column , text=column)

    # Insert data into the treeview
    for row in excel_reader.get_rows(path):
        treeview.insert("" , tk.END , values=row)
    tree_y_Scrollbar.config(command=treeview.yview)
    tree_x_Scrollbar.config(command=treeview.xview)

    frames_list.append(treeview)

    # Create entry fields for data insertion
    entries = [customtkinter.CTkEntry(insert_frame, placeholder_text=column) for column in excel_reader.get_columns(path)]
    for entry in entries:
        entry.pack(side="top", pady=10)

    # Add button to insert data
    insert_button = ttk.Button(insert_frame ,text='Insert Data' , command=lambda:insert_data([entry.get() for entry in entries] , path , frames_list))
    insert_button.pack(pady=10)

# Call the first_page function
first_page()
# Start the main event loop
app.mainloop()
