import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
from tkinter import messagebox

target_column_names = [
    "Participating Corporation",
    "Date",
    "First Name",
    "Last Name",
    "Phone Number",
    "Email",
    "Address",
    "City",
    "State",
    "Postal",
    "Platform",
    "Comment",
    "Transaction ID",
    "Donation Frequency",
    "Donation Amount",
    "Matched Amount",
    "Total",
    "Fee",
    "Net",
    "Monthly Total",
    "Yearly Total",
    "Campaigns",
    "Source Name",
    "Charge Time"
]

platforms = [
    "Benevity",
    "Facebook",
    "KCEGF",
    "PayPal",
    "PPGF",
    "UW-CFD",
]

root = tk.Tk()
root.title = "PSF Spreadsheet Converter"

selected_filetype = tk.StringVar()
to_convert_filename = ""

def BenevityHandler(converted_filename):
    df = pd.read_csv(to_convert_filename, skiprows=11, skipfooter=4, engine="python")
    new_df = pd.DataFrame(columns=target_column_names)
    new_df["Participating Corporation"] = df.iloc[:, 0].values
    new_df["Date"] = df.iloc[:, 2].values
    new_df["First Name"] = df.iloc[:, 3].values
    new_df["Last Name"] = df.iloc[:, 4].values
    #new_df["Phone Number"] = df.iloc[:, 0].values
    new_df["Email"] = df.iloc[:, 5].values
    new_df["Address"] = df.iloc[:, 6].values
    new_df["City"] = df.iloc[:, 7].values
    new_df["State"] = df.iloc[:, 8].values
    new_df["Postal"] = df.iloc[:, 9].values
    new_df["Platform"] = "Benevity"
    new_df["Comment"] = df.iloc[:, 11].values
    new_df["Transaction ID"] = df.iloc[:, 12].values
    new_df["Donation Frequency"] = df.iloc[:, 13].values
    new_df["Donation Amount"] = df.iloc[:, 18].values
    new_df["Matched Amount"] = df.iloc[:, 19].values
    new_df["Total"] = df.iloc[:, 18].values + df.iloc[:, 19].values
    new_df["Fee"] = df.iloc[:, 20].values + df.iloc[:, 21].values
    new_df["Net"] = (df.iloc[:, 18].values + df.iloc[:, 19].values) - (df.iloc[:, 20].values + df.iloc[:, 21].values)
    #new_df["Monthly Total"] = df.iloc[:, 0].values
    #new_df["Yearly Total"] = df.iloc[:, 0].values
    new_df["Campaigns"] = df.iloc[:, 10].values
    #new_df["Source Name"] = df.iloc[:, 0].values
    #new_df["Charge Time"] = df.iloc[:, 0].values
    new_df.to_csv(converted_filename)

def FacebookHandler(converted_filename):
    df = pd.read_csv(to_convert_filename, engine="python")
    new_df = pd.DataFrame(columns=target_column_names)
    #new_df["Participating Corporation"] = df.iloc[:, 0].values
    new_df["Date"] = df.iloc[:, 10].values
    new_df["First Name"] = df.iloc[:, 11].values
    new_df["Last Name"] = df.iloc[:, 12].values
    #new_df["Phone Number"] = df.iloc[:, 0].values
    new_df["Email"] = df.iloc[:, 13].values
    #new_df["Address"] = df.iloc[:, 6].values
    #new_df["City"] = df.iloc[:, 7].values
    #new_df["State"] = df.iloc[:, 8].values
    #new_df["Postal"] = df.iloc[:, 9].values
    new_df["Platform"] = "Facebook"
    #new_df["Comment"] = df.iloc[:, 11].values
    new_df["Transaction ID"] = df.iloc[:, 0].values
    #new_df["Donation Frequency"] = df.iloc[:, 13].values
    new_df["Donation Amount"] = df.iloc[:, 2].values
    #new_df["Matched Amount"] = df.iloc[:, 19].values
    new_df["Total"] = df.iloc[:, 2].values
    new_df["Fee"] = df.iloc[:, 3].values
    new_df["Net"] = df.iloc[:, 4].values
    #new_df["Monthly Total"] = df.iloc[:, 0].values
    #new_df["Yearly Total"] = df.iloc[:, 0].values
    new_df["Campaigns"] = df.iloc[:, 15].values
    new_df["Source Name"] = df.iloc[:, 16].values
    new_df["Charge Time"] = df.iloc[:, 23].values
    new_df.to_csv(converted_filename)

def KcegfHandler(converted_filename):
    df = pd.read_excel(to_convert_filename)
    new_df = pd.DataFrame(columns=target_column_names)
    #new_df["Participating Corporation"] = df.iloc[:, 0].values
    new_df["Date"] = "Quarter " + df.iloc[:, 0].astype(str)
    new_df["First Name"] = df.iloc[:, 4].values
    new_df["Last Name"] = df.iloc[:, 5].values
    #new_df["Phone Number"] = df.iloc[:, 0].values
    #new_df["Email"] = df.iloc[:, 5].values
    new_df["Address"] = df.iloc[:, 9].values
    new_df["City"] = df.iloc[:, 11].values
    new_df["State"] = df.iloc[:, 12].values
    new_df["Postal"] = df.iloc[:, 13].values
    new_df["Platform"] = "KCEGF"
    #new_df["Comment"] = df.iloc[:, 11].values
    new_df["Transaction ID"] = df.iloc[:, 3].values
    #new_df["Donation Frequency"] = df.iloc[:, 13].values
    new_df["Donation Amount"] = df.iloc[:, 6].values
    #new_df["Matched Amount"] = df.iloc[:, 19].values
    new_df["Total"] = df.iloc[:, 6].values
    #new_df["Fee"] = df.iloc[:, 20].values + df.iloc[:, 21].values
    new_df["Net"] = df.iloc[:, 6].values
    #new_df["Monthly Total"] = df.iloc[:, 0].values
    #new_df["Yearly Total"] = df.iloc[:, 0].values
    #new_df["Campaigns"] = df.iloc[:, 10].values
    #new_df["Source Name"] = df.iloc[:, 0].values
    #new_df["Charge Time"] = df.iloc[:, 0].values
    new_df.to_csv(converted_filename)

def PaypalHandler(converted_filename):
    df = pd.read_csv(to_convert_filename, engine="python")
    new_df = pd.DataFrame(columns=target_column_names)
    #new_df["Participating Corporation"] = df.iloc[:, 0].values
    new_df["Date"] = df.iloc[:, 0].values
    new_df["First Name"] = df.iloc[:, 3].values
    #new_df["Last Name"] = df.iloc[:, 4].values
    new_df["Phone Number"] = df.iloc[:, 36].values
    new_df["Email"] = df.iloc[:, 10].values
    new_df["Address"] = df.iloc[:, 30].values
    new_df["City"] = df.iloc[:, 32].values
    new_df["State"] = df.iloc[:, 33].values
    new_df["Postal"] = df.iloc[:, 34].values
    new_df["Platform"] = "PayPal"
    new_df["Comment"] = df.iloc[:, 37].values
    new_df["Transaction ID"] = df.iloc[:, 12].values
    #new_df["Donation Frequency"] = df.iloc[:, 13].values
    new_df["Donation Amount"] = df.iloc[:, 7].values
    #new_df["Matched Amount"] = df.iloc[:, 19].values
    new_df["Total"] = df.iloc[:, 7].values
    new_df["Fee"] = df.iloc[:, 8].values
    new_df["Net"] = df.iloc[:, 9].values
    #new_df["Monthly Total"] = df.iloc[:, 0].values
    #new_df["Yearly Total"] = df.iloc[:, 0].values
    #new_df["Campaigns"] = df.iloc[:, 10].values
    #new_df["Source Name"] = df.iloc[:, 0].values
    new_df["Charge Time"] = df.iloc[:, 1].values
    new_df.to_csv(converted_filename)

def PpgfHandler(converted_filename):
    df = pd.read_csv(to_convert_filename, engine="python")
    new_df = pd.DataFrame(columns=target_column_names)
    #new_df["Participating Corporation"] = df.iloc[:, 0].values
    new_df["Date"] = df.iloc[:, 0].values
    new_df["First Name"] = df.iloc[:, 1].values
    #new_df["Last Name"] = df.iloc[:, 4].values
    #new_df["Phone Number"] = df.iloc[:, 36].values
    new_df["Email"] = df.iloc[:, 2].values
    #new_df["Address"] = df.iloc[:, 30].values
    #new_df["City"] = df.iloc[:, 32].values
    #new_df["State"] = df.iloc[:, 33].values
    #new_df["Postal"] = df.iloc[:, 34].values
    new_df["Platform"] = "PPGF"
    #new_df["Comment"] = df.iloc[:, 37].values
    #new_df["Transaction ID"] = df.iloc[:, 12].values
    #new_df["Donation Frequency"] = df.iloc[:, 13].values
    new_df["Donation Amount"] = df.iloc[:, 7].values
    #new_df["Matched Amount"] = df.iloc[:, 19].values
    new_df["Total"] = df.iloc[:, 7].values
    #new_df["Fee"] = df.iloc[:, 8].values
    new_df["Net"] = df.iloc[:, 7].values
    #new_df["Monthly Total"] = df.iloc[:, 0].values
    #new_df["Yearly Total"] = df.iloc[:, 0].values
    new_df["Campaigns"] = df.iloc[:, 4].values
    new_df["Source Name"] = df.iloc[:, 3].values
    #new_df["Charge Time"] = df.iloc[:, 1].values
    new_df.to_csv(converted_filename)

def UwCfdHandler(converted_filename):
    df = pd.read_excel(to_convert_filename, sheetname = "Full Year", skiprows=2, skipfooter=5)
    new_df = pd.DataFrame(columns=target_column_names)
    #new_df["Participating Corporation"] = df.iloc[:, 0].values
    new_df["Date"] = df.iloc[:, 18].astype(str)
    new_df["First Name"] = df.iloc[:, 7].values
    new_df["Last Name"] = df.iloc[:, 8].values
    #new_df["Phone Number"] = df.iloc[:, 0].values
    new_df["Email"] = df.iloc[:, 9].values
    new_df["Address"] = df.iloc[:, 10].values
    new_df["City"] = df.iloc[:, 11].values
    new_df["State"] = df.iloc[:, 12].values
    new_df["Postal"] = df.iloc[:, 13].values
    new_df["Platform"] = "UW-CFD"
    #new_df["Comment"] = df.iloc[:, 11].values
    new_df["Transaction ID"] = df.iloc[:, 17].values
    #new_df["Donation Frequency"] = df.iloc[:, 13].values
    new_df["Donation Amount"] = df.iloc[:, 6].values
    #new_df["Matched Amount"] = df.iloc[:, 19].values
    new_df["Total"] = df.iloc[:, 6].values
    #new_df["Fee"] = df.iloc[:, 20].values + df.iloc[:, 21].values
    new_df["Net"] = df.iloc[:, 6].values
    #new_df["Monthly Total"] = df.iloc[:, 0].values
    #new_df["Yearly Total"] = df.iloc[:, 0].values
    #new_df["Campaigns"] = df.iloc[:, 10].values
    #new_df["Source Name"] = df.iloc[:, 0].values
    #new_df["Charge Time"] = df.iloc[:, 0].values
    new_df.to_csv(converted_filename)

def SelectFileType():
    print(selected_filetype.get(), "selected")
    
def SelectFileToConvert():
    global to_convert_filename
    to_convert_filename = filedialog.askopenfilename(filetypes=[("All files", "*.*"), ("Excel file","*.xlsx"),("Excel file 97-2003","*.xls"), ("Comma-separated values file", "*.csv")])
    print(to_convert_filename, "selected for conversion")
    
def ConvertAndSaveFile():
    converted_filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Comma-separated values file", "*.csv")])
    if (selected_filetype.get() == "Benevity"):
        BenevityHandler(converted_filename)
    elif (selected_filetype.get() == "Facebook"):
        FacebookHandler(converted_filename)
    elif (selected_filetype.get() == "KCEGF"):
        KcegfHandler(converted_filename)
    elif (selected_filetype.get() == "PayPal"):
        PaypalHandler(converted_filename)
    elif (selected_filetype.get() == "PPGF"):
        PpgfHandler(converted_filename)
    elif (selected_filetype.get() == "UW-CFD"):
        UwCfdHandler(converted_filename)

tk.Label(root, 
         text="""Choose the spreadsheet type: """,
         justify = tk.LEFT,
         padx = 20).pack()

# Add file type options
for platform in platforms:
    tk.Radiobutton(root, 
                  text=platform,
                  indicatoron = 0,
                  width = 20,
                  padx = 20, 
                  variable=selected_filetype, 
                  command=SelectFileType,
                  value=platform).pack(anchor=tk.W)

# Add file selection
tk.Button(root,
          text="Select file to convert",
          command = SelectFileToConvert).pack(pady=20)

# Add save file button
tk.Button(root,
          text="Save converted file as",
          command = ConvertAndSaveFile).pack(pady=(0, 10))

root.mainloop()

