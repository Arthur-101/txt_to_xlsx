import re

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

# Extracting Result
result = re.search(result_pattern, input_string).group()

# Extracting Marks and Grades
marks_grades = re.findall(marks_grades_pattern, input_string)

# Separate individual subject codes from subject_codes_string
subject_codes = re.findall(r"\d{3}", subject_codes_string)

print("Roll:", roll)
print("Gender:", gender)
print("Name:", name.strip())
print("Subject Codes:", subject_codes)
print("Result:", result)
print("Marks and Grades:", marks_grades)
