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
marks_grades_pattern = r"(\d{3}\s\s[A-D][1-9])"

# Extracting Roll, Gender, and Name
roll, gender, name = re.search(roll_gender_name_pattern, input_string).groups()

# Extracting Subject Codes
subject_codes_string = re.search(subject_codes_pattern, input_string).group()

# Extracting Result
result = re.search(result_pattern, input_string).group()

# Extracting Marks and Grades
marks_grades = re.findall(marks_grades_pattern, input_string)

# Separate individual subject codes from subject_codes_string
subject_codes = re.findall(r"\d{3}", subject_codes_string)

# Create the dataframe
data = {
    'Roll': [roll, np.NaN],
    'Gender': [gender, np.NaN],
    'Name': [name.strip(), np.NaN],
    'Sub_1': [subject_codes[0], marks_grades[0]],
    'Sub_2': [subject_codes[1], marks_grades[1]],
    'Sub_3': [subject_codes[2], marks_grades[2]],
    'Sub_4': [subject_codes[3], marks_grades[3]],
    'Sub_5': [subject_codes[4], marks_grades[4]],
    'Sub_6': [subject_codes[5], marks_grades[5]],
    'Result': [result, np.NaN],
}
df = pd.DataFrame(data)

print(df)
