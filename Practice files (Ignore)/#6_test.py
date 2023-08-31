import re
import pandas as pd

def extract_data(file_name):
  with open(file_name, 'r') as f:
    lines = f.readlines()

  data = []
  for line in lines:
    match = re.match(r'^(\d{10})\s+(\w)\s+(.*)$', line)
    if match:
      roll_number, gender, name = match.groups()
      subjects = []
      for subject in line.split(' ')[5:]:
        grade = subject[-1]
        subject_code = subject[:-1]
        subjects.append((subject_code, grade))
      data.append((roll_number, gender, name,) + tuple(subjects))

  df = pd.DataFrame(data, columns=['Roll', 'Gender', 'Name', 'Sub_1', 'Grade_1', 'Sub_2', 'Grade_2', 'Sub_3', 'Grade_3', 'Sub_4', 'Grade_4', 'Sub_5', 'Grade_5', 'Sub_6', 'Grade_6', 'Result', 'Compart_Sub'])

  # Extract the data from the file using regular expressions
  for i in range(len(df)):
    for col in ['Sub_1', 'Sub_2', 'Sub_3', 'Sub_4', 'Sub_5', 'Sub_6']:
      grade = df.loc[i, col][-1]
      subject_code = df.loc[i, col][:-1]
      df.loc[i, col] = (subject_code, grade)

  # Set the result and compartment subject columns
  for i in range(len(df)):
    result = 'PASS'
    compartment_sub = []
    for col in ['Sub_1', 'Sub_2', 'Sub_3', 'Sub_4', 'Sub_5', 'Sub_6']:
      grade = df.loc[i, col][1]
      if grade != 'A1' and grade != 'B1':
        result = 'COMP'
        compartment_sub.append(df.loc[i, col][0])
    df.loc[i, 'Result'] = result
    df.loc[i, 'Compart_Sub'] = compartment_sub

  return df

df = extract_data('65027 XII.txt')

print(df.to_string())
