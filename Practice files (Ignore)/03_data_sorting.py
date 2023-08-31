import re
import pandas as pd

# Read the text file
with open('65027 Xii.txt', 'r') as file:
    text = file.read()

# Define patterns for extracting Date, School code and name, and Region
date_pattern = r"DATE:- (\d{2}/\d{2}/\d{4})"
school_pattern = r"SCHOOL : - (\d+) (.+)"
region_pattern = r"REGION: (.+)"

# Find matches for Date, School code and name, and Region
date_match = re.search(date_pattern, text)
school_match = re.search(school_pattern, text)
region_match = re.search(region_pattern, text)

date = date_match.group(1) if date_match else ""
school_code = school_match.group(1) if school_match else ""
school_name = school_match.group(2) if school_match else ""
region = region_match.group(1) if region_match else ""

# Define the pattern to extract data for each candidate
candidate_pattern = r"(\d+)\s+(\w)\s+(\w.+?)\s+(\d{3})\s+(\d{3})\s+(\d{3})\s+(\d{3})\s+(\d{3})\s+(\d{3})\s+(\w \w \w)\s+(\w+)\s*"

# Find all matches in the text using the candidate pattern
matches = re.findall(candidate_pattern, text)

# Initialize empty lists to store data
roll_numbers = []
genders = []
names = []
subjects = []
grades = []
results = []
comp_subjects = []

# Process each match and extract data points
for match in matches:
    roll_number = match[0]
    gender = match[1]
    name = match[2].strip()
    subjects_list = match[3:9]
    grades_list = match[9:15]
    result = match[15]
    comp_subject = match[16].strip()

    roll_numbers.append(roll_number)
    genders.append(gender)
    names.append(name)
    subjects.append(subjects_list)
    grades.append(grades_list)
    results.append(result)
    comp_subjects.append(comp_subject)

# Create a pandas dataframe
df = pd.DataFrame({
    'Roll No.': roll_numbers,
    'Gender': genders,
    'Student Name': names,
    'Subject 1': [s[0] for s in subjects],
    'Grade 1': [g[0] for g in grades],
    'Subject 2': [s[1] for s in subjects],
    'Grade 2': [g[1] for g in grades],
    'Subject 3': [s[2] for s in subjects],
    'Grade 3': [g[2] for g in grades],
    'Subject 4': [s[3] for s in subjects],
    'Grade 4': [g[3] for g in grades],
    'Subject 5': [s[4] for s in subjects],
    'Grade 5': [g[4] for g in grades],
    'Subject 6': [s[5] for s in subjects],
    'Grade 6': [g[5] for g in grades],
    'Result': results,
    'Comp Subject': comp_subjects
})

# Add additional information (Date, School Code, School Name, Region) to the dataframe
df['Date'] = date
df['School Code'] = school_code
df['School Name'] = school_name
df['Region'] = region

# Display the dataframe
print(df)
