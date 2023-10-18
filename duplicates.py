import pandas as pd
from fuzzywuzzy import fuzz
from datetime import datetime


def check_similarity(name1, name2):
    # Check if both names are strings
    if isinstance(name1, str) and isinstance(name2, str):
        components1 = set(name1.split())
        components2 = set(name2.split())

        if len(components1) > 1 and len(components2) > 1:
            # Count common substrings
            common_substrings = len(components1.intersection(components2))

            # Check if at least two identical substrings
            if common_substrings >= 2:
                return True

        # Use fuzz ratio to check similarity (including handling typos)
        similarity_ratio = fuzz.ratio(name1, name2)

        # You can adjust the threshold based on your needs
        fuzzy_similarity = similarity_ratio > 85
        # similarity_ratio > 80  # Adjust the threshold as needed

        # Regular similarity check
        return fuzzy_similarity
    else:
        return False  # Skip non-string values

# Load the Excel file into a DataFrame
file_path = 'C:\\Users\\USER\\Desktop\\projects\\hamal\\data\\mondayCopy.xlsx'
df = pd.read_excel(file_path)

# Get the names from column A
names = df.iloc[:100, 0].tolist()

# Dictionary to store similar names
similar_names_dict = {}

# Iterate through names and check for similarity
for i in range(len(names)):
    for j in range(i + 1, len(names)):
        if check_similarity(names[i], names[j]):
            # Add similar names to the dictionary
            if names[i] not in similar_names_dict:
                similar_names_dict[names[i]] = set()
            similar_names_dict[names[i]].add(names[j])

# Create a DataFrame for the similar name groups
output_df = pd.DataFrame(list(similar_names_dict.items()), columns=['MainName', 'SimilarNames'])

# Sort the DataFrame alphabetically based on 'MainName'
output_df = output_df.sort_values(by='MainName')

# Create a new Excel file with the sorted DataFrame
current_time = datetime.now().time()
formatted_time = current_time.strftime('%H_%M_%S')
output_file_path = f'C:\\Users\\USER\\Desktop\\projects\\hamal\\file_grouped{formatted_time}.xlsx'
output_df.to_excel(output_file_path, index=False)

print(f'Similar names grouped and saved to: {output_file_path}')