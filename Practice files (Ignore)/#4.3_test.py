import re
import pandas as pd
import numpy as np

input_string = "22614342   M ABHIMANYU KUMAR                                     301     302     041     042     043     048      B1 A2 B2    PASS\n                                  076  B2 080  B1 051  D1 063  C1 052  D2 083  B1"

# Pattern to extract Roll, Gender, and Name
roll_gender_name_pattern = r"(\d+)\s+(\w)\s+([A-Z ]+)"

# Pattern to extract Subject Codes
subject_codes_pattern = r"(\d{3}\s+){5}\d{3}"

# Pattern to extract Result
result_pattern = r"\b(PASS|FAIL|COMP)\b"

# Pattern to extract Marks and Grades
marks_grades_pattern = r"(\d{3}\s+[A-D][1-9])"

# Extracting Roll, Gender, and Name
roll, gender, name = re.search(roll_gender_name_pattern, input_string).groups()

# Extracting Subject Codes
subject_codes_string = re.search(subject_codes_pattern, input_string).group()
subject_codes = re.findall(r"\d{3}", subject_codes_string)

# Extracting Result
result = re.search(result_pattern, input_string).group()

# Extracting Marks and Grades
marks_grades = re.findall(marks_grades_pattern, input_string)

# Create a DataFrame with column names
df_data = {
    'Roll': [roll, np.nan],
    'Gender': [gender, np.nan],
    'Name': [name.strip(), np.nan]
}

# Add Subject Codes (Sub_1, Sub_2, ..., Sub_6) to the DataFrame
for i, code in enumerate(subject_codes):
    df_data[f'Sub_{i+1}'] = [code, np.nan]

# Add Result to the DataFrame
df_data['Result'] = [result, np.nan]

# Create the DataFrame
df = pd.DataFrame(df_data)

# Add the second row for Marks and Grades
df.loc[1] = [np.nan] * df.shape[1]
for i, mark_grade in enumerate(marks_grades):
    df[f'Sub_{i+1}'][1] = mark_grade

print(df)
