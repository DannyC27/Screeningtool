import os
import re
import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkinter import filedialog
import sys
import pyodbc

EQ_DB_user_name = 'Geosyntec\Daniel.Chu'

conEQ = "Driver={ODBC Driver 17 for SQL Server}; " + "Server=SQL-PM02; " + "Database=prjEQuISTraining; " + "Trusted_Connection=Yes; " + "Uid=" + EQ_DB_user_name + ";"  # + "Pwd="+ EQ_DB_pwd + ";"

query = "select * from PFAS_Chemicals"

# --- Global Variables ---
df1 = None
df2 = PFAS_df
cas_column1 = 0
chemical_name_column = 0


conn = pyodbc.connect(conEQ)
PFAS_df = pd.read_sql(query, con=conn)

# --- Statistics Tracking ---
comparisons_made = 0
definite_matches = 0
potential_matches = 0
screened_out = 0

# --- Function Definitions ---

def update_cas_column(cas_var, file_index):
    global cas_column1
    selected_column = cas_var.get()
    cas_column1 = df1.columns.get_loc(selected_column)

def update_chemical_name_column(chemical_name_var, file_index):
    global chemical_name_column
    selected_column = chemical_name_var.get()
    chemical_name_column = df1.columns.get_loc(selected_column)

# Read Excel files
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

# Perform PSS for keywords can add more later
def partial_string_search(item, keywords):
    if item is None:
        return 0  # Handle empty cells
    return 0.5 if any(keyword in item.lower() for keyword in keywords) else 0

# Clean CAS numbers (remove all non-numeric characters and dashes)
def clean_cas_number(cas):
    if cas is None:
        return ""
    cas = str(cas)
    return re.sub(r"[^0-9-]", "", cas)

# Definite Match Screening
def definite_match(cas_numbers, chemical_names, df2, cas_column2, definite_yes):
    global comparisons_made, definite_matches
    for i, cas in enumerate(cas_numbers):
        comparisons_made += 1
        # No need to clean CAS here, we want original form
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"Checking definite match for CAS: {cas}\n")
        console_output.config(state=tk.DISABLED)
        if cas in df2[df2.columns[cas_column2]].tolist():
            definite_yes.append(f"{cas}: {chemical_names[i]}")  # Keep original names
            definite_matches += 1
            update_stats()

# Full Name Comparison (Add to definite yes if found)
def full_name_comparison(cas_numbers, chemical_names, df2, chemical_name_column, definite_yes, remaining_items, chemical_name_column2):
    global comparisons_made, definite_matches
    for i, cas in enumerate(cas_numbers):
        comparisons_made += 1
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"Checking full name comparison for CAS: {cas}\n")
        console_output.config(state=tk.DISABLED)
        if cas in remaining_items:
            # No need to clean CAS here, we want original form
            for j, name2 in enumerate(df2[df2.columns[chemical_name_column2]].tolist()):
                if chemical_names[i] == name2: # Compare original names
                    definite_yes.append(f"{cas}: {chemical_names[i]}")  # Keep original names
                    remaining_items.remove(cas)  # Remove from remaining items
                    definite_matches += 1
                    update_stats()
                    console_output.config(state=tk.NORMAL)
                    console_output.insert(tk.END, f"  Found full name match for CAS: {cas}\n")
                    console_output.config(state=tk.DISABLED)
                    break  # Move to the next CAS number

# Maybe Match Screening
def maybe_match(cas_numbers, chemical_names, remaining_items, maybe_matches, no_matches, keywords):
    global comparisons_made, potential_matches
    for i, cas in enumerate(cas_numbers):  # Iterate through the original cas_numbers
        comparisons_made += 1
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"Checking maybe match for CAS: {cas}\n")
        console_output.config(state=tk.DISABLED)
        if cas in remaining_items:
            result = partial_string_search(chemical_names[i], keywords)
            # No need to clean CAS here, we want original form
            if result == 0.5:  # Check for any match
                maybe_matches.append(f"{cas}: {chemical_names[i]}")  # Keep original names
                remaining_items.remove(cas)  # Remove from remaining_items if it's a maybe match
                potential_matches += 1
                update_stats()
                console_output.config(state=tk.NORMAL)
                console_output.insert(tk.END, f"  Found maybe match for CAS: {cas}\n")
                console_output.config(state=tk.DISABLED)
            else:
                no_matches.append(f"{clean_cas_number(cas)}: {chemical_names[i]}")  # Clean CAS for no match

# No Match Screening (for remaining items)
def no_match(cas_numbers, chemical_names, remaining_items, no_matches):
    global comparisons_made, screened_out
    for i, cas in enumerate(cas_numbers):
        comparisons_made += 1
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"Checking no match for CAS: {cas}\n")
        console_output.config(state=tk.DISABLED)
        if cas in remaining_items:
            no_matches.append(f"{clean_cas_number(cas)}: {chemical_names[i]}")  # Clean CAS for no match
            remaining_items.remove(cas)  # Remove the CAS from remaining items after processing
            screened_out += 1
            update_stats()
            console_output.config(state=tk.NORMAL)
            console_output.insert(tk.END, f"  Found no match for CAS: {cas}\n")
            console_output.config(state=tk.DISABLED)

# Create data.
def compare_excel_files(df1, df2, cas_column1, chemical_name_column, cas_column2, chemical_name_column2):
    global comparisons_made, definite_matches, potential_matches, screened_out
    # Reset statistics before each comparison
    comparisons_made = 0
    definite_matches = 0
    potential_matches = 0
    screened_out = 0
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
    definite_match(cas_numbers, chemical_names, df2, cas_column2, definite_yes)

    # Get items that didn't get a definite yes
    remaining_items = [cas for i, cas in enumerate(cas_numbers) if cas not in df2[df2.columns[cas_column2]].tolist()]

    full_name_comparison(cas_numbers, chemical_names, df2, chemical_name_column, definite_yes, remaining_items, chemical_name_column2)

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

def update_dropdown(df, cas_var, chemical_name_var, cas_dropdown, chemical_name_dropdown):
    cas_var.set(df.columns[0])
    if chemical_name_var is not None:
        chemical_name_var.set(df.columns[0])
    cas_dropdown['values'] = list(df.columns)
    if chemical_name_dropdown is not None:
        chemical_name_dropdown['values'] = list(df.columns)

def compare():
    global cas_column1, chemical_name_column
    try:
        compare_excel_files(df1, df2, cas_column1, chemical_name_column, 0, 1)  # REPLACE CAS COLUMN NUMBER WITH 0 AND NAME WITH 1
    except Exception as e:
        console_output.config(state=tk.NORMAL)
        console_output.insert(tk.END, f"An error occurred: {e}\n")
        console_output.config(state=tk.DISABLED)

def update_stats():
    stats_label.config(text=f"Comparisons Made: {comparisons_made}\n"
                            f"Definite Matches: {definite_matches}\n"
                            f"Potential Matches: {potential_matches}\n"
                            f"Screened Out: {screened_out}")

# --- GUI Setup ---
root = tk.Tk()
root.title("PFAS reporting rule tool - daniel.chu@geosyntec.com")

# --- Stats Frame ---
stats_frame = tk.Frame(root, width=200, bg="#f0f0f0")  # Light gray background
stats_frame.pack(side=tk.LEFT, fill=tk.Y, padx=20)

stats_label = tk.Label(stats_frame, text="Comparisons Made: 0\n"
                                        "Definite Matches: 0\n"
                                        "Potential Matches: 0\n"
                                        "Screened Out: 0", justify="left", font=("Arial", 12), bg="#f0f0f0")
stats_label.pack(side=tk.TOP, pady=20)

# --- File Import Frame ---
file_frame = tk.Frame(root)
file_frame.pack(pady=20)

file_path1 = tk.StringVar(file_frame)

# File 1 import
file1_label = tk.Label(file_frame, text="File 1:")
file1_label.grid(row=0, column=0, padx=10)
file1_entry = tk.Entry(file_frame, width=50, textvariable=file_path1)
file1_entry.grid(row=0, column=1, padx=10)
import_button1 = tk.Button(file_frame, text="Import File 1", command=import_file1, width=15, bg="#4CAF50", fg="white")  # Green button
import_button1.grid(row=0, column=2, pady=10)

# --- Column Selection Frame ---
column_frame = tk.Frame(root)
column_frame.pack()

# File 1 column selection
cas_column_label1 = tk.Label(column_frame, text="CAS Column (File 1):")
cas_column_label1.grid(row=0, column=0, padx=10)
cas_column_var1 = tk.StringVar(column_frame)
cas_column_dropdown1 = ttk.Combobox(column_frame, textvariable=cas_column_var1, state="readonly", width=20)
cas_column_dropdown1.bind("<<ComboboxSelected>>", lambda event: update_cas_column(cas_column_var1, 0))
cas_column_dropdown1.grid(row=0, column=1, padx=10)

chemical_name_label1 = tk.Label(column_frame, text="Chemical Name Column (File 1):")
chemical_name_label1.grid(row=1, column=0, padx=10)
chemical_name_var1 = tk.StringVar(column_frame)
chemical_name_dropdown1 = ttk.Combobox(column_frame, textvariable=chemical_name_var1, state="readonly", width=20)
chemical_name_dropdown1.bind("<<ComboboxSelected>>", lambda event: update_chemical_name_column(chemical_name_var1, 0))
chemical_name_dropdown1.grid(row=1, column=1, padx=10)

# --- Buttons ---
compare_button = tk.Button(root, text="Compare", command=compare, state=tk.DISABLED, width=15, bg="#4CAF50", fg="white")  # Green button
compare_button.pack(pady=10)

download_button = tk.Button(root, text="Download Results", command=download_results, state=tk.DISABLED, width=15, bg="#2196F3", fg="white")  # Blue button
download_button.pack(pady=20)

# --- Console Output Frame ---
console_frame = tk.Frame(root)
console_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)  # Expand the console frame

console_output = tk.Text(console_frame, state=tk.DISABLED, wrap=tk.WORD, font=("Arial", 11), height=20)
console_output.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)  # Expand the console output

# --- Scrollbar ---
console_scrollbar = tk.Scrollbar(console_frame, command=console_output.yview, width=20)  # Make the scrollbar wider
console_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
console_output.config(yscrollcommand=console_scrollbar.set)

root.mainloop()