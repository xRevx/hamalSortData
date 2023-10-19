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
            for comp1 in components1:
                for comp2 in components2:
                    if fuzz.ratio(comp1, comp2) > 75:
                        return True
        # Use fuzz ratio to check similarity (including handling typos)
        similarity_ratio = fuzz.ratio(name1, name2)
    
        # You can adjust the threshold based on your needs
        fuzzy_similarity = similarity_ratio > 90
        #similarity_ratio > 80  # Adjust the threshold as needed

        # Regular similarity check
        
        return fuzzy_similarity
    else:
        return False  # Skip non-string values

# Load the Excel file into a DataFrame
all_names_path = 'C:\\Users\\USER\\Desktop\\projects\\hamal\\data\\mondayCopy.xlsx'#change to your excal with all of the names
df = pd.read_excel(all_names_path)
names_to_check_path = 'C:\\Users\\USER\\Desktop\\projects\\hamal\\data\\19.10.xlsx'#change to your excal with all of the people u want to filter with
data = pd.read_excel(names_to_check_path)

# Get the names from column A
names = data.iloc[:, 0].tolist()
allNames = df.iloc[:, 0].tolist()


# List to store similar name pairs
similar_names = []
count_iterations = 0
# Iterate through names and check for similarity
for i in range(len(names)):
    for j in range(i + 1, len(allNames)):
        count_iterations+=1
        if check_similarity(names[i], allNames[j]):
            similar_names.append((names[i], allNames[j], i, j))
print(count_iterations)


# Create a DataFrame for the similar name pairs
output_df = pd.DataFrame(similar_names, columns=['Name1', 'Name2', 'name1rowID', 'name2RowID'])

# Sort the DataFrame alphabetically based on 'Name1'

# Create a new Excel file with the sorted DataFrame
current_time = datetime.now().time()
formatted_time = current_time.strftime('%H_%M_%S')
output_file_path = f'C:\\Users\\USER\\Desktop\\projects\\hamal\\data\\{formatted_time}.xlsx' #put ur output file 
output_df.to_excel(output_file_path, index=False)

print(f'Similar names saved and sorted to: {output_file_path}')
