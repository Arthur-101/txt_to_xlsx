import pandas as pd
import numpy as np

# Your original dataframe
data = {
    "Roll": [22614342, 22614343, 22614344, 22614345, 22614346, 22614347],
    "Gender": ["M", "F", "F", "M", "M", "M"],
    "Name": ["ABHIMANYU KUMAR", "AKANKSHA PRIYA", "AKANKSHI KUMARI", "AKEE KUMAR", "AMAN KUMAR", "AMISH RANJAN"],
    "Sub_1": [301, 301, 301, 301, 301, 301],
    "Marks_1": [76, 86, 84, 54, 79, 78],
    "grade_1": ["B2", "A2", "B1", "D1", "B2", "B2"],
    "Sub_2": [302, 302, 322, 48, 302, 48],
    "Marks_2": [80, 76, 98, 86, 74, 85],
    "grade_2": ["B1", "B2", "A1", "A2", "B2", "B1"],
    "Sub_3": [41, 41, 41, 41, 41, 41],
    "Marks_3": [51, 75, 95, 55, 53, 56],
    "grade_3": ["D1", "B1", "A1", "C1", "C2", "C1"],
    "Sub_4": [42, 42, 42, 42, 42, 42],
    "Marks_4": [63, 79, 95, 66, 62, 61],
    "grade_4": ["C1", "A2", "A1", "B2", "C1", "C2"],
    "Sub_5": [43, 43, 43, 43, 43, 43],
    "Marks_5": [52, 74, 95, 59, 65, 66],
    "grade_5": ["D2", "B1", "A1", "C2", "C1", "B2"],
    "Sub_6": [48, 48, 48, 65, 48, 65],
    "Marks_6": [83, 94, 89, 73, 85, 80],
    "grade_6": ["B1", "A1", "A2", "C2", "B1", "C1"],
}

df = pd.DataFrame(data)

# Extract unique subject codes
subject_codes = df[['Sub_1', 'Sub_2', 'Sub_3', 'Sub_4', 'Sub_5', 'Sub_6']].values.flatten()
subject_codes = np.unique(subject_codes)

# Create a new dataframe
new_columns = ['Roll', 'Gender', 'Name'] + subject_codes.tolist()
new_df = pd.DataFrame(columns=new_columns)

# Iterate through rows and populate the new dataframe
for _, row in df.iterrows():
    new_row = [row['Roll'], row['Gender'], row['Name']]
    for code in subject_codes:
        matching_index = np.where(row[['Sub_1', 'Sub_2', 'Sub_3', 'Sub_4', 'Sub_5', 'Sub_6']] == code)[0]
        if matching_index.size > 0:
            new_row.append(row['Marks_' + str(matching_index[0] + 1)])
        else:
            new_row.append(np.nan)
    new_df.loc[len(new_df)] = new_row

print(new_df)
