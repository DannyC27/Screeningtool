import os
import re
import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkinter import filedialog
import sys
import pyodbc

#EQ_DB_user_name = 'Geosyntec\Daniel.Chu'

#conEQ = "Driver={ODBC Driver 17 for SQL Server}; " + "Server=SQL-PM02; " + "Database=prjEQuISTraining; " + "Trusted_Connection=Yes; " + "Uid=" + EQ_DB_user_name + ";"  # + "Pwd="+ EQ_DB_pwd + ";"

#query = "select * from PFAS_Chemicals"

#conn = pyodbc.connect(conEQ)
#PFAS_df = pd.read_sql(query, con=conn)

# Initialize global variables
df1 = None
df2 = None
cas_column1 = 0
chemical_name_column = 0
cas_column2 = 0

# Function to update CAS column selection
def update_cas_column(cas_var, file_index):
    global cas_column1, cas_column2
    selected_column = cas_var.get()
    if file_index == 0:
        cas_column1 = df1.columns.get_loc(selected_column)
    else:
        cas_column2 = df2.columns.get_loc(selected_column)

# Function to update chemical name column selection
def update_chemical_name_column(chemical_name_var, file_index):
    global chemical_name_column
    selected_column = chemical_name_var.get()
    chemical_name_column = df1.columns.get_loc(selected_column)

# Function to read files
def read_excel(file_path):
    console_output.config(state=tk.NORMAL)
    console_output.insert(tk.END, f"Trying to read: {file_path}\n")
    console_output.config(state=tk.DISABLED)
    try:
        return pd.read_excel(file_path)
    except FileNotFoundError:
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"File not found: {file_path}\n")
        console_output.config(state=tk.DISABLED)
        return None

# Function to perform partial string search
def partial_string_search(item, keywords):
    if item is None:
        return 0  # Handle empty cells
    return 0.5 if any(keyword in item.lower() for keyword in keywords) else 0

# Function to clean chemical names
def clean_chemical_name(name):
    if name is None:
        return ""
    # Convert the name to a string before using re.sub
    name = str(name)
    return re.sub(r"[^a-zA-Z\s]", "", name).strip()

# Function to clean CAS numbers
def clean_cas_number(cas):
    if cas is None:
        return ""
    cas = str(cas)
    return re.sub(r"[^0-9-]", "", cas)

# Function for definite match screening
def definite_match(cas_numbers, chemical_names, df2, cas_column2, definite_yes):
    for i, cas in enumerate(cas_numbers):
        cleaned_cas = clean_cas_number(cas)  # Clean the CAS number
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"Checking definite match for CAS: {cleaned_cas}\n")
        console_output.config(state=tk.DISABLED)
        if cleaned_cas in df2[df2.columns[cas_column2]].tolist():
            definite_yes.append(f"{cleaned_cas}: {clean_chemical_name(chemical_names[i])}")

# Function for full name comparison
def full_name_comparison(cas_numbers, chemical_names, df2, chemical_name_column, definite_yes, remaining_items):
    for i, cas in enumerate(cas_numbers):
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"Checking full name comparison for CAS: {cas}\n")
        console_output.config(state=tk.DISABLED)
        if cas in remaining_items:
            cleaned_name1 = clean_chemical_name(chemical_names[i])
            for j, name2 in enumerate(df2[df2.columns[chemical_name_column]].tolist()):
                cleaned_name2 = clean_chemical_name(name2)
                if cleaned_name1 == cleaned_name2:
                    definite_yes.append(f"{clean_cas_number(cas)}: {cleaned_name1}")
                    remaining_items.remove(cas)  # Remove from remaining items
                    console_output.config(state=tk.NORMAL)
                    console_output.insert(tk.END, f"  Found full name match for CAS: {cas}\n")
                    console_output.config(state=tk.DISABLED)
                    break  # Move to the next CAS number

# Function for maybe match screening
def maybe_match(cas_numbers, chemical_names, remaining_items, maybe_matches, no_matches, keywords):
    for i, cas in enumerate(cas_numbers):  # Iterate through the original cas_numbers
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"Checking maybe match for CAS: {cas}\n")
        console_output.config(state=tk.DISABLED)
        if cas in remaining_items:
            # Check for empty names or partial matches
            if chemical_names[i] == "" or partial_string_search(chemical_names[i], keywords) == 0.5:
                maybe_matches.append(f"{cas}: {clean_chemical_name(chemical_names[i])}")  # Output original CAS
                remaining_items.remove(cas)  # Remove from remaining_items if it's a maybe match
                console_output.config(state=tk.NORMAL)
                console_output.insert(tk.END, f"  Found maybe match for CAS: {cas}\n")
                console_output.config(state=tk.DISABLED)
            else:
                no_matches.append(f"{clean_cas_number(cas)}: {clean_chemical_name(chemical_names[i])}")

# Function for no match screening
def no_match(cas_numbers, chemical_names, remaining_items, no_matches):
    for i, cas in enumerate(cas_numbers):
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"Checking no match for CAS: {cas}\n")
        console_output.config(state=tk.DISABLED)
        if cas in remaining_items:
            no_matches.append(f"{clean_cas_number(cas)}: {clean_chemical_name(chemical_names[i])}")
            remaining_items.remove(cas)  # Remove the CAS from remaining items after processing
            console_output.config(state=tk.NORMAL)
            console_output.insert(tk.END, f"  Found no match for CAS: {cas}\n")
            console_output.config(state=tk.DISABLED)

# Function to perform the comparison based on selected screening method
def compare_excel_files(df1, df2, cas_column1, chemical_name_column, cas_column2, screening_method):
    # create the buckets (large arraylists)
    definite_yes = []
    maybe_matches = []
    no_matches = []

    # Get CAS Numbers and Chemical Names
    cas_numbers = df1[df1.columns[cas_column1]].tolist()
    chemical_names = df1[df1.columns[chemical_name_column]].tolist()

    # Keywords for Partial String Search
    keywords = ["fluor", "fluo"]

    # Screening Levels
    if screening_method == "CAS Only":
        definite_match(cas_numbers, chemical_names, df2, cas_column2, definite_yes)

        # Get items that didn't get a definite yes
        remaining_items = [cas for i, cas in enumerate(cas_numbers) if clean_cas_number(cas) not in df2[df2.columns[cas_column2]].tolist()]

        no_match(cas_numbers, chemical_names, remaining_items, no_matches)

    elif screening_method == "Name Only":
        full_name_comparison(cas_numbers, chemical_names, df2, chemical_name_column, definite_yes, cas_numbers.copy())

        maybe_match(cas_numbers, chemical_names, cas_numbers.copy(), maybe_matches, no_matches, keywords)

        # No match screening isn't needed for Name Only as all items are checked

    elif screening_method == "Full Screening":
        definite_match(cas_numbers, chemical_names, df2, cas_column2, definite_yes)

        # Get items that didn't get a definite yes
        remaining_items = [cas for i, cas in enumerate(cas_numbers) if clean_cas_number(cas) not in df2[df2.columns[cas_column2]].tolist()]

        full_name_comparison(cas_numbers, chemical_names, df2, chemical_name_column, definite_yes, remaining_items)

        maybe_match(cas_numbers, chemical_names, remaining_items, maybe_matches, no_matches, keywords)

        no_match(cas_numbers, chemical_names, remaining_items, no_matches)

    # Output
    console_output.config(state=tk.NORMAL)
    console_output.insert(tk.END, "\nResults:\n")
    console_output.insert(tk.END, "Definite Yes (CAS Number Match or Full Name Match):\n")
    console_output.config(state=tk.DISABLED)
    for match in definite_yes:
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"{match}\n")
        console_output.config(state=tk.DISABLED)

    console_output.config(state=tk.NORMAL)
    console_output.insert(tk.END, "\nMaybe (Contains Keyword indicators or CAS number uncertain):\n")
    console_output.config(state=tk.DISABLED)
    for match in maybe_matches:
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"{match}\n")
        console_output.config(state=tk.DISABLED)

    console_output.config(state=tk.NORMAL)
    console_output.insert(tk.END, "\nNo Match:\n")
    console_output.config(state=tk.DISABLED)
    for match in no_matches:
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"{match}\n")
        console_output.config(state=tk.DISABLED)

    # Create a DataFrame for each match type
    definite_yes_df = pd.DataFrame([item.split(": ") for item in definite_yes],
        columns=['CAS', 'Chemical Name'])
    definite_yes_df['Match Score'] = 1
    maybe_matches_df = pd.DataFrame(
        [item.split(": ") for item in maybe_matches],
        columns=['CAS', 'Chemical Name'])
    maybe_matches_df['Match Score'] = 0.5
    no_matches_df = pd.DataFrame([item.split(": ") for item in no_matches],
        columns=['CAS', 'Chemical Name'])
    no_matches_df['Match Score'] = 0

    # Save the DataFrames to an Excel file each dataframe is a new sheet
    with pd.ExcelWriter('results.xlsx') as writer:
        definite_yes_df.to_excel(writer,
                                 sheet_name='Definite Matches',
                                 index=False)
        maybe_matches_df.to_excel(writer,
                                  sheet_name='Maybe Matches',
                                  index=False)
        no_matches_df.to_excel(writer, sheet_name='No Matches', index=False)

    console_output.config(state=tk.NORMAL)
    console_output.insert(tk.END, "\nResults saved to 'results.xlsx'.\n")
    console_output.config(state=tk.DISABLED)
    download_button.config(state=tk.NORMAL)  # Enable download button after saving

def download_results():
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if save_path:
        os.rename("results.xlsx", save_path)  # Rename to the chosen save path

def import_file1():
    global df1, cas_column1, chemical_name_column
    file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        file_path1.set(file_path)
        df1 = read_excel(file_path)
        if df1 is not None:
            update_dropdown(df1, cas_column_var1, chemical_name_var1, cas_column_dropdown1, chemical_name_dropdown1)
            if df2 is not None:  # Enable Compare button if both files are loaded
                compare_button.config(state=tk.NORMAL)

def import_file2():
   global df2, cas_column2
   file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
   if file_path:
        file_path2.set(file_path)
        df2 = read_excel(file_path)
        if df2 is not None:
            update_dropdown(df2, cas_column_var2, None, cas_column_dropdown2, None)
            if df1 is not None:  # Enable Compare button if both files are loaded
                compare_button.config(state=tk.NORMAL)

def update_dropdown(df, cas_var, chemical_name_var, cas_dropdown, chemical_name_dropdown):
    cas_var.set(df.columns[0])
    if chemical_name_var is not None:
        chemical_name_var.set(df.columns[0])
    cas_dropdown['values'] = list(df.columns)
    if chemical_name_dropdown is not None:
        chemical_name_dropdown['values'] = list(df.columns)

# Function to handle compare button click
def compare():
    global cas_column1, chemical_name_column, cas_column2
    screening_method = screening_var.get()
    compare_excel_files(df1, df2, cas_column1, chemical_name_column, cas_column2, screening_method)

# Initialize GUI
root = tk.Tk()
root.title("PFAS Screening tool - Daniel.Chu@geosyntec.com")

# File Import Frame
file_frame = tk.Frame(root)
file_frame.pack(pady=20)

file_path1 = tk.StringVar(file_frame)
file_path2 = tk.StringVar(file_frame)

# File 1 import
file1_label = tk.Label(file_frame, text="File 1:")
file1_label.grid(row=0, column=0)
file1_entry = tk.Entry(file_frame, width=50, textvariable=file_path1)
file1_entry.grid(row=0, column=1)
import_button1 = tk.Button(file_frame, text="Import File 1", command=import_file1)
import_button1.grid(row=0, column=2, pady=10)

# File 2 import
file2_label = tk.Label(file_frame, text="Master Database")
file2_label.grid(row=1, column=0)
file2_entry = tk.Entry(file_frame, width=50, textvariable=file_path2)
file2_entry.grid(row=1, column=1)
import_button2 = tk.Button(file_frame, text="Import File 2", command=import_file2)
import_button2.grid(row=1, column=2, pady=10)

# Column Selection Frame
column_frame = tk.Frame(root)
column_frame.pack()

# File 1 column selection
cas_column_label1 = tk.Label(column_frame, text="CAS Column (File 1):")
cas_column_label1.grid(row=0, column=0, padx=10)
cas_column_var1 = tk.StringVar(column_frame)
cas_column_dropdown1 = ttk.Combobox(column_frame, textvariable=cas_column_var1, state="readonly")
cas_column_dropdown1.bind("<<ComboboxSelected>>", lambda event: update_cas_column(cas_column_var1, 0))
cas_column_dropdown1.grid(row=0, column=1, padx=10)

chemical_name_label1 = tk.Label(column_frame, text="Chemical Name Column (File 1):")
chemical_name_label1.grid(row=1, column=0, padx=10)
chemical_name_var1 = tk.StringVar(column_frame)
chemical_name_dropdown1 = ttk.Combobox(column_frame, textvariable=chemical_name_var1, state="readonly")
chemical_name_dropdown1.bind("<<ComboboxSelected>>", lambda event: update_chemical_name_column(chemical_name_var1, 0))
chemical_name_dropdown1.grid(row=1, column=1, padx=10)

# File 2 column selection
cas_column_label2 = tk.Label(column_frame, text="CAS Column (File 2):")
cas_column_label2.grid(row=0, column=2, padx=10)
cas_column_var2 = tk.StringVar(column_frame)
cas_column_dropdown2 = ttk.Combobox(column_frame, textvariable=cas_column_var2, state="readonly")
cas_column_dropdown2.bind("<<ComboboxSelected>>", lambda event: update_cas_column(cas_column_var2, 1))
cas_column_dropdown2.grid(row=0, column=3, padx=10)

# Screening Method Selection Frame
screening_frame = tk.Frame(root)
screening_frame.pack()

screening_var = tk.StringVar(screening_frame)
screening_var.set("Full Screening")  # Default screening method

# Create radio buttons for screening methods
screening_options = [
    ("Full Screening", "Full Screening"),
    ("CAS Only", "CAS Only"),
    ("Name Only", "Name Only")
]

for text, value in screening_options:
    screening_radio = tk.Radiobutton(screening_frame, text=text, variable=screening_var, value=value)
    screening_radio.pack(anchor=tk.W)

# Compare Button
compare_button = tk.Button(root, text="Compare", command=compare, state=tk.DISABLED)
compare_button.pack(pady=10)

# Download Button
download_button = tk.Button(root, text="Download Results", command=download_results, state=tk.DISABLED)
download_button.pack(pady=20)

# Console Output Frame
console_frame = tk.Frame(root)
console_frame.pack(side=tk.RIGHT, fill=tk.Y)

console_output = tk.Text(console_frame, state=tk.DISABLED, wrap=tk.WORD)
console_output.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

# Scrollbar
console_scrollbar = tk.Scrollbar(console_frame, command=console_output.yview)
console_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
console_output.config(yscrollcommand=console_scrollbar.set)

root.mainloop()