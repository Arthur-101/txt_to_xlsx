import re

text = "22614342   M ABHIMANYU KUMAR                                     301     302     041     042     043     048      B1 A2 B2    PASS                       \n                                                                 076  B2 080  B1 051  D1 063  C1 052  D2 083  B1"

# Regular expressions to capture different pieces of data
roll_number_pattern = r"(\d+)"
gender_pattern = r"\s+(\w)\s+"
name_pattern = r"([A-Z\s]+)"
subject_codes_pattern = r"\s+(\d{3})\s+(\d{3})\s+(\d{3})\s+(\d{3})\s+(\d{3})\s+(\d{3})\s+"
result_pattern = r"([A-Z]+)\s+"

# Marks and Grades patterns
marks_pattern = r"(\d{3})\s+"
grades_pattern = r"([A-Z0-9]+)\s+"

# Extracting data using regular expressions
roll_number_match = re.search(roll_number_pattern, text)
gender_match = re.search(gender_pattern, text)
name_match = re.search(name_pattern, text)
subject_codes_match = re.search(subject_codes_pattern, text)
result_match = re.search(result_pattern, text)

marks_matches = re.findall(marks_pattern, text)
grades_matches = re.findall(grades_pattern, text)

# Assigning captured data to separate variables
roll_number = roll_number_match.group(1) if roll_number_match else None
gender = gender_match.group(1) if gender_match else None
name = name_match.group(1).strip() if name_match else None

subject_codes = (
    subject_codes_match.groups() if subject_codes_match else (None, None, None, None, None, None)
)

result = result_match.group(1) if result_match else None

# Since the text contains six marks and six grades, we can use a loop to extract them into separate variables
marks = marks_matches[:6]
grades = grades_matches[:6]

# Printing the extracted data
print("Roll Number:", roll_number)
print("Gender:", gender)
print("Name:", name)
print("Subject Codes:", subject_codes)
print("Result:", result)
print("Marks:", marks)
print("Grades:", grades)
