import os

import re

import pandas as pd


# Read Excel files
def read_excel(file_path):
    print(f"Trying to read: {file_path}")  # Debug for reading
    try:
        return pd.read_excel(file_path)
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return None


# Perform PSS for keywords can add more later
def partial_string_search(item, keywords):
    if item is None:
        return 0  # Handle empty cells
    return 0.5 if any(keyword in item.lower() for keyword in keywords) else 0


# Slice Chemical Names (Remove numbers and special characters)
def clean_chemical_name(name):
    if name is None:
        return ""
    # Convert the name to a string before using re.sub
    name = str(name)
    return re.sub(r"[^a-zA-Z\s]", "", name).strip()


# Clean CAS numbers (remove all non-numeric characters and dashes)
def clean_cas_number(cas):
    if cas is None:
        return ""
    cas = str(cas)
    return re.sub(r"[^0-9-]", "", cas)


# Definite Match Screening
def definite_match(cas_numbers, chemical_names, df2, cas_column2,
                   definite_yes):
    for i, cas in enumerate(cas_numbers):
        cleaned_cas = clean_cas_number(cas)  # Clean the CAS number
        print(f"Checking definite match for CAS: {cleaned_cas}")  # Debug print
        if cleaned_cas in df2[df2.columns[cas_column2]].tolist():
            print(f"  Found Definite match for CAS: {cas}")
            definite_yes.append(
                f"{cleaned_cas}: {clean_chemical_name(chemical_names[i])}")
        else:
            print(f"  No definite match found for CAS: {cas}")


# Full Name Comparison (Add to definite yes if found)
def full_name_comparison(cas_numbers, chemical_names, df2,
                         chemical_name_column, definite_yes, remaining_items):
    for i, cas in enumerate(cas_numbers):
        print(f"Checking full name comparison for CAS: {cas}")  # Debug print
        if cas in remaining_items:
            cleaned_name1 = clean_chemical_name(chemical_names[i])
            for j, name2 in enumerate(
                    df2[df2.columns[chemical_name_column]].tolist()):
                cleaned_name2 = clean_chemical_name(name2)
                if cleaned_name1 == cleaned_name2:
                    definite_yes.append(
                        f"{clean_cas_number(cas)}: {cleaned_name1}")
                    remaining_items.remove(cas)  # Remove from remaining items
                    print(f"  Found full name match for CAS: {cas}"
                          )  # Debug print
                    break  # Move to the next CAS number


# Maybe Match Screening
def maybe_match(cas_numbers, chemical_names, remaining_items, maybe_matches,
                keywords):
    for i, cas in enumerate(
            cas_numbers):  # Iterate through the original cas_numbers
        print(f"Checking maybe match for CAS: {cas}")  # Debug print
        if cas in remaining_items:
            result = partial_string_search(chemical_names[i], keywords)
            cleaned_cas = clean_cas_number(cas)
            # Convert 'cas' to a string for the check
            cas_str = str(cas)
            if result == 0.5:  # Check for any match
                maybe_matches.append(
                    f"{cas}: {clean_chemical_name(chemical_names[i])}"
                )  # Output original CAS
                remaining_items.remove(
                    cas)  # Remove from remaining_items if it's a maybe match
                print(f"  Found maybe potential for CAS: {cas}")  # Debug print


# No Match Screening (for remaining items)
def no_match(cas_numbers, chemical_names, remaining_items, no_matches):
    for i, cas in enumerate(cas_numbers):
        print(f"Checking no match for CAS: {cas}")  # Debug print
        if cas in remaining_items:
            no_matches.append(
                f"{clean_cas_number(cas)}: {clean_chemical_name(chemical_names[i])}"
            ) 
            print(f"  Found no match for CAS: {cas}")  
            remaining_items.remove(
                cas)  # Remove the CAS from remaining items after processing
            # Debug print


# Create data. THIS IS THE MAIN GUTS WHICH CALLS UPON ALL OTHER FUNCTIONS AND OUTPUTS DATA
def compare_excel_files(file_path1, file_path2):
    df1 = read_excel(file_path1)
    df2 = read_excel(file_path2)

    if df1 is None or df2 is None:
        return None

    # Get column names and display options to the user as a list
    print(f"\nColumns in '{file_path1}':")
    for i, column in enumerate(df1.columns):
        print(f"{i+1}. {column}")

    cas_column1 = int(
        input(
            f"Enter the number corresponding to the CAS numbers column in '{file_path1}': "
        )) - 1
    chemical_name_column = int(
        input(
            f"Enter the number corresponding to the chemical names column in '{file_path1}': "
        )) - 1

    print(f"\nColumns in '{file_path2}':")
    for i, column in enumerate(df2.columns):
        print(f"{i+1}. {column}")

    cas_column2 = int(
        input(
            f"Enter the number corresponding to the CAS numbers column in '{file_path2}': "
        )) - 1

    # create the buckets (large arraylists)
    definite_yes = []
    maybe_matches = []
    no_matches = []

    # Get CAS Numbers and Chemical Names
    cas_numbers = df1[df1.columns[cas_column1]].tolist()
    chemical_names = df1[df1.columns[chemical_name_column]].tolist()

    # Keywords for Partial String Search
    keywords = [
        "fluor", "fluo", "PFOA", "Perfluoro", "perfluoro", "perfluor",
        "PFEESA", "HFPO-DA", "NFDHA", "PFOS", "PFUnA", "NMeFOSAA", "PFPeA",
        "PFPeS", "6:2 FTS", "NEtFOSAA", "FBSA", "PFHxA", "PFDoA", "PFOA",
        "PFDA", "PFDS", "PFHxS", "PFBA", "PFBS", "PFHpA", "PFHpS", "PFNA",
        "PFTeA", "PFMPA", "8:2 FTS", "FHxSA", "PFPrS", "PFNS", "PFTriA",
        "9Cl-PF3ONS", "FOSA", "4:2 FTS", "11Cl-PF3OUdS", "PFECHS", "PFMBA",
        "ADONA", "PFOA+PFOS"
    ]

    # Screening Levels
    definite_match(cas_numbers, chemical_names, df2, cas_column2, definite_yes)

    # Get items that didn't get a definite yes
    remaining_items = [
        cas for i, cas in enumerate(cas_numbers)
        if clean_cas_number(cas) not in df2[df2.columns[cas_column2]].tolist()
    ]

    full_name_comparison(cas_numbers, chemical_names, df2,
                         chemical_name_column, definite_yes, remaining_items)

    maybe_match(cas_numbers, chemical_names, remaining_items, maybe_matches,
                keywords)

    no_match(cas_numbers, chemical_names, remaining_items, no_matches)

    # Output
    print("\nResults:")
    print("Definite Yes (CAS Number Match or Full Name Match):")
    for match in definite_yes:
        print(match)
    print("\nMaybe (Contains Keyword indicators or CAS number uncertain):")
    for match in maybe_matches:
        print(match)
    print("\nNo Match:")
    for match in no_matches:
        print(match)

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

    print("Results saved to 'results.xlsx'.")


# Main function
if __name__ == "__main__":
    print(f"Current working directory: {os.getcwd()}")  # Print the directory
    file_path1 = input("Enter the path to the first Excel file: ")
    file_path2 = input("Enter the path to the second Excel file: ")
    results = compare_excel_files(file_path1, file_path2)

    if results:
        print("Results generated.")
    else:
        print("Debugline: End of Generation")
