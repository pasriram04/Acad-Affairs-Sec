import pandas as pd
import numpy as np

file_path = r"C:\Users\hp\Downloads\Academic Options Form (Responses).xlsx"
final_file_path = r"C:\Users\hp\Downloads\Results.xlsx"

name_list = [
    "Minor in Aerospace Engineering", "Minor in Artificial Intelligence", "Minor in Biomedical Engineering", "Minor in Climate Change", "Minor in Chemical Engineering", 
    "Minor in Civil Engineering", "Minor in Computer Science and Engineering", "Minor in Creative Arts", "Minor in Design", "Minor in Economics",
    "Minor in Electrical Engineering", "Minor in Entrepreneurship", "Minor in Engineering Physics", "Minor in Materials Science and Metallurgical Engineering", "Minor in Mathematics", 
    "Minor in Mechanical Engineering", "Double Major in Chemical Engineering", "Double Major in Civil Engineering", "Double Major in Computer Science and Engineering", "Double Major in Engineering Physics", 
    "Double Major in Entrepreneurship", "Double Major in Electrical Engineering", "Double Major in Materials Science and Metallurgical Engineering", "Double Major in Mathematics", "Double Major in Mechanical Engineering"
]

# Define the list with the number of seats, set to 5 for all 25 elements
'''seats_list = [5, 5, 5, 5, 5,
              5, 5, 5, 5, 5,
              5, 5, 5, 5, 5,
              5, 5, 5, 5, 5,
              5, 5, 5, 5, 5]'''
seats_list = [2] * 25

# Guidelines
# 1. The column containing CGPA must be named "CGPA" and nothing else
# 2. There must be 12 priority columns
# 3. Use same spellings in the form as code. Form link: https://docs.google.com/forms/d/1tewxJwSxIsTXiF_KwwzQaPvD7VsFfRHEw9MuNs80Yfs/edit

df = pd.read_excel(file_path)
df.columns = df.columns.str.strip() # Trim leading and trailing spaces from column names

# Check if 'CGPA' column exists after trimming spaces
if 'CGPA' in df.columns:
    # Sort the DataFrame by the 'CGPA' column
    df = df.sort_values(by='CGPA', ascending=False)
else:
    print("'CGPA' column not found in the Excel file")
    exit()

priority_columns = [f'Priority {i}' for i in range(1, 13)]
existing_priority_columns = [col for col in df.columns if any(f'Priority {i}' in col for i in range(1, 13))]

# Ensure there are exactly 12 columns
if len(existing_priority_columns) != 12:
    print(f"Expected 12 priority columns, but found {len(existing_priority_columns)}: {existing_priority_columns}")
    exit()
else:
    # Extract data from these columns and store it in a 2-dimensional array
    priority_data = df[existing_priority_columns].values

allotted = []
students = []

for index, row in enumerate(priority_data):
    flag = 0
    parent_department = df.iloc[index]['Name of Parent Program/Department']
    for string in row:
        if pd.notna(string):
            if string.count(', ') > 1 or string.lower().count('Major') > 1 or parent_department in string:
                print(f"Index: {index}, More than one majors or two minors or major/minor in parent department: {string}")
                continue
            if ', ' in string:
                string_1, string_2 = [s.strip() for s in string.split(', ', 1)]
                if (string_1 in name_list) and (string_2 in name_list):
                    index_in_name_list_1 = name_list.index(string_1)
                    index_in_name_list_2 = name_list.index(string_2)
                    if (seats_list[index_in_name_list_1] != 0) and (seats_list[index_in_name_list_2] != 0):
                        allotted.append(string_1 + ", " + string_2)
                        seats_list[index_in_name_list_1] -= 1
                        seats_list[index_in_name_list_2] -= 1
                        flag = 1
                        students.append(df.iloc[index]['Name of Student'])
                        break  # Exit the inner loop
                else:
                    print(string_1 + string_2)
            elif pd.notna(string):
                if string in name_list:
                    index_in_name_list = name_list.index(string)
                    if seats_list[index_in_name_list] != 0:
                        allotted.append(string)
                        seats_list[index_in_name_list] -= 1
                        flag = 1
                        students.append(df.iloc[index]['Name of Student'])
                        break  # Exit the inner loop
                else:
                    print(string)
    if(flag == 0):
        allotted.append("No allocation")
        students.append(df.iloc[index]['Name of Student'])

# Create a DataFrame for the allocation result including student names
allocation_df = pd.DataFrame({'Name of Student': students, 'Allotted': allotted})

# Save the DataFrame to an Excel file
allocation_df.to_excel(final_file_path, index=False)
print(f"Allotted sorted by 'CGPA' with student names and saved to {final_file_path}")
