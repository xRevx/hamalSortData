import pandas as pd
from fuzzywuzzy import fuzz
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import re
import math

def is_nan(value):
    return math.isnan(value)

def not_equal(name,name2):
    if(check_not_null(name) and check_not_null(name2) and not is_nan(name) and not(name2)):
        if name != name2:
            return True
    return False    

def check_not_null(name):
    if(name == None or name == ""):
        return True
    return False


def check_similarity(name1, name2):
    if isinstance(name1, str) and isinstance(name2, str):
        if name1 == name2:
            return True
        components1 = set(name1.split())
        components2 = set(name2.split())

        if len(components1) > 1 and len(components2) > 1:
            # Count common substrings
            common_substrings = len(components1.intersection(components2))

            # Check if at least two identical substrings
            if common_substrings >= 2:
                return True
        # Use fuzz ratio to check similarity (including handling typos)
        similarity_ratio = fuzz.ratio(name1,name2)
    
        

        fuzzy_similarity = similarity_ratio > 90

        return fuzzy_similarity
    else:
        return False  # Skip non-string valuesng values or not meeting similarity conditions

def extract_digits(phone_number):
    # Convert to string if it's a float
    phone_str = str(phone_number)

    # Remove non-digit characters from the phone number
    digits_only = re.sub(r'\D', '', phone_str)
    
    # Check if the phone number starts with '972-' and remove it
    if digits_only.startswith('972') and len(digits_only) > 3:
        digits_only = digits_only[3:]
    elif digits_only.startswith('05') and len(digits_only) > 2:
        digits_only = digits_only[2:]
    elif digits_only.startswith('972-') and len(digits_only) > 4:
        digits_only = digits_only[4:]
    elif digits_only.startswith('054-') and len(digits_only) > 4:
        digits_only = digits_only[4:]
    
    return digits_only

all_names_path = 'C:\\Users\\USER\\Desktop\\projects\\hamal\\data\\Copy.xlsx'  # change to your Excel with all names
df_all_names = pd.read_excel(all_names_path)

# Load the Excel file with names to check
names_to_check_path = 'C:\\Users\\USER\\Desktop\\projects\\hamal\\data\\19.10.xlsx'  # change to your Excel with names to check
df_names_to_check = pd.read_excel(names_to_check_path)

# Get the names from column A
names = df_names_to_check.iloc[:, 0].tolist()

# Create a Workbook
output_wb = Workbook()

# Create a worksheet
output_ws = output_wb.active

# List to store rows that satisfy the similarity condition
similar_rows = []
similar_rows.append(["***********************"])
copy_rows = []
last_rows = []
last_rows.append(["***********************"])

phone_miss = ""
phone_maybe = ""
id_missing = ""
id_maybe = ""
# Iterate through names and check for similarity
for i in range(len(names)):

    phone_maybe = df_names_to_check.iloc[i, 2]
    phone_maybe = extract_digits(phone_maybe)
    id_maybe = df_names_to_check.iloc[i, 1]
    id_maybe = str(id_maybe)
    id_maybe = re.sub(r'\D', '', id_maybe)



    # Append the original name in red
    for j in range(len(df_all_names)):
        if check_similarity(names[i], df_all_names.iloc[j, 0]):
            phone_miss = df_all_names.iloc[i, 4]
            phone_miss = extract_digits(phone_miss)
            id_missing = df_all_names.iloc[i, 6]
            id_missing = str(id_missing)
            id_missing = re.sub(r'\D', '', id_missing)

            print(phone_miss)
            print(phone_maybe)
            print(id_missing)
            print(id_maybe)
            original_name = f'***{names[i]}***'

            if id_missing == id_maybe or phone_maybe == phone_miss:
                copy_rows.append([original_name])
                print(1)
                # Append the entire row to the list
                copy_rows.append(df_all_names.iloc[j, :].tolist())  # Stop checking for this name once a match is found
            elif not_equal(id_missing, id_maybe) and not_equal(phone_maybe,phone_miss):
                print(2)
                last_rows.append([original_name])
                # Append the entire row to the list
                last_rows.append(df_all_names.iloc[j, :].tolist())  # Stop checking for this name once a match is found
            else:    
                similar_rows.append([original_name])
                print(3)

                # Append the entire row to the list
                similar_rows.append(df_all_names.iloc[j, :].tolist())  # Stop checking for this name once a match is found


for row in copy_rows:
    output_ws.append(row)
for row in similar_rows:
    output_ws.append(row)
for row in last_rows:
    output_ws.append(row)

current_time = datetime.now().time()
formatted_time = current_time.strftime('%H_%M_%S')
output_file_path = f'C:\\Users\\USER\\Desktop\\projects\\hamal\\data\\similar_rows_{formatted_time}.xlsx'
output_wb.save(output_file_path)

print(f'Similar rows copied and saved to: {output_file_path}')
