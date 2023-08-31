import re

current_student_grades = "076 B2 080 B1 051 D1 063 C1 052 D2 083 B1"
# Pattern to extract Marks and Grades for each subject
marks_grades_pattern = r"(\d{3})\s+([A-D][1-9])"
# Extracting Marks and Grades for each subject
marks_grades = re.findall(marks_grades_pattern, current_student_grades)

# Separate marks and grades
subject_marks = [subject[0] for subject in marks_grades]
subject_grades = [subject[1] for subject in marks_grades]

print("Subject Marks:", subject_marks)
print("Subject Grades:", subject_grades)
