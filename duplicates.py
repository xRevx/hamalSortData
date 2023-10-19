import pandas as pd
from fuzzywuzzy import fuzz
from datetime import datetime


def check_similarity(name1, name2):
    # Check if both names are strings
    if isinstance(name1, str) and isinstance(name2, str):
        if name1 == name2:
            return True

        components1 = set(name1.split())
        components2 = set(name2.split())

        # Check if at least two identical substrings
        common_substrings = len(components1.intersection(components2))
        if common_substrings >= 2:
            return True
        else:
            # Check fuzzywuzzy similarity between non-matching components
            for comp1 in components1 - components2:  # Only non-matching components from components1
                for comp2 in components2 - components1:  # Only non-matching components from components2
                    if fuzz.ratio(comp1, comp2) > 85:
                        return True


    return False  # Skip non-string values or not meeting similarity conditions


# Load the Excel file into a DataFrame
all_names_path = 'C:\\Users\\USER\\Desktop\\projects\\hamal\\data\\mondayCopy.xlsx'  # change to your Excel with all names
df_all_names = pd.read_excel(all_names_path)

# Load the Excel file with names to check
names_to_check_path = 'C:\\Users\\USER\\Desktop\\projects\\hamal\\data\\19.10.xlsx'  # change to your Excel with names to check
df_names_to_check = pd.read_excel(names_to_check_path)

# Get the names from column A
names = df_names_to_check.iloc[:, 0].tolist()

# List to store rows that satisfy the similarity condition
similar_rows = []

# Iterate through names and check for similarity
for i in range(len(names)):
    for j in range(len(df_all_names)):
        if check_similarity(names[i], df_all_names.iloc[j, 0]):
            # Append the entire row to the list
            similar_rows.append([f'<font color="red">{names[i]}</font>'])
            similar_rows.append(df_all_names.iloc[j, :].tolist())
            break  # Stop checking for this name once a match is found

# Create a DataFrame from the list of similar rows
output_df = pd.DataFrame(similar_rows, columns=df_all_names.columns)

# Create a new Excel file with the DataFrame
current_time = datetime.now().time()
formatted_time = current_time.strftime('%H_%M_%S')
output_file_path = f'C:\\Users\\USER\\Desktop\\projects\\hamal\\data\\similar_rows_{formatted_time}.xlsx'
output_df.to_excel(output_file_path, index=False)

print(f'Similar rows copied and saved to: {output_file_path}')
