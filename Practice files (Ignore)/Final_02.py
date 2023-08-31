import re
import pandas as pd
import numpy as np

# Reading the text file:-
with open('65027 XII.txt','r') as file:
    reader = file.read()

# Defining the patterns to get 'Date', 'School Code', 'School Name' and Region' :-
date_pattern = r"DATE:- (\d{2}/\d{2}/\d{4})"
school_pattern = r"SCHOOL : - (\d+) (.+)"
region_pattern = r"REGION:\s+(\S+)"

# Finding the matches for 'Date', 'School code', 'School name', and 'Region' :-
date_match = re.search(date_pattern, reader)
school_match = re.search(school_pattern, reader)
region_match = re.search(region_pattern, reader)

# Giving variable name to all the data:-
date = date_match.group(1) if date_match else ""
school_code = school_match.group(1) if school_match else ""
school_name = school_match.group(2) if school_match else ""
region = region_match.group(1) if region_match else ""

# Pattern to match the unwanted text (from the top of the string to the line starting with "SCHOOL")
unwanted_pattern = r'DATE:.*?\n.*?-----.*?\n\nSCHOOL.*?\n'

# Remove unwanted text using re.sub
input_string_cleaned = re.sub(unwanted_pattern, '', reader, flags=re.DOTALL)

# Pattern to extract Roll, Gender, and Name
roll_gender_name_pattern = r"(\d+)\s+(\w)\s+([A-Z ]+)"

# Pattern to extract Subject Codes
subject_codes_pattern = r"(\d{3}\s+){5}(\d{3})?"

# Pattern to extract Result
result_pattern = r"\b(PASS|FAIL|COMP)\b"

# Pattern to extract Marks and Grades for each subject
marks_grades_pattern = r"\d{3}\s+[A-Z]\d?"

# Initialize an empty list to store each student's data as a dictionary
students_data = []

# Initialize variables to keep track of student information
current_student_info = None
current_student_grades = None

# Split the input_string by newline characters
lines = reader.strip().split('\n')

# Function to process and add student data to students_data list
def add_student_data(roll, gender, name, subject_codes, result, marks_grades):
    marks = []
    grades = []
    for mark_grade in marks_grades:
        mark, grade = mark_grade.split()
        marks.append(int(mark))
        grades.append(grade)
    
    if len(subject_codes) < 6 and len(marks) < 6 and len(grades) < 6:
        subject_codes.append(np.NaN)
        marks.append(np.NaN)
        grades.append(np.NaN)

    row_data = {
        'Roll': roll,
        'Gender': gender,
        'Name': name.strip(),
        'Sub_1': subject_codes[0],
        'Marks_1': marks[0],
        'grade_1': grades[0],
        'Sub_2': subject_codes[1],
        'Marks_2': marks[1],
        'grade_2': grades[1],
        'Sub_3': subject_codes[2],
        'Marks_3': marks[2],
        'grade_3': grades[2],
        'Sub_4': subject_codes[3],
        'Marks_4': marks[3],
        'grade_4': grades[3],
        'Sub_5': subject_codes[4],
        'Marks_5': marks[4],
        'grade_5': grades[4],
        'Sub_6': subject_codes[5],
        'Marks_6': marks[5],
        'grade_6': grades[5],
        'Result': result
    }
    
    students_data.append(row_data)
    
# Iterate through each line to process student data
for line in lines:
    line = line.strip()
    # Check if the line contains Roll, Gender, and Name
    if re.match(roll_gender_name_pattern, line):
        current_student_info = line
    # Check if the line contains Marks and Grades
    elif re.match(marks_grades_pattern, line):
        current_student_grades = line
        # Extracting Roll, Gender, and Name
        roll, gender, name = re.search(roll_gender_name_pattern, current_student_info).groups()

        # Extracting Subject Codes and Separating individual subject codes
        subject_codes_string = re.search(subject_codes_pattern, current_student_info).group()
        subject_codes = re.findall(r"\d{3}", subject_codes_string)

        # Extracting Result
        result = re.search(result_pattern, current_student_info).group()

        # Extracting Marks and Grades for each subject
        marks_grades = re.findall(marks_grades_pattern, current_student_grades)

        # Add student data to students_data list
        add_student_data(roll, gender, name, subject_codes, result, marks_grades)

# Create the final DataFrame using the list of student dictionaries
df = pd.DataFrame(students_data)
df = df.fillna("")

# Reset the index of the DataFrame
df.reset_index(drop=True, inplace=True)

# Print the DataFrame
print(df)
