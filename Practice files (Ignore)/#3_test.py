import re
import pandas as pd

# Read the content of the text file
with open('input.txt', 'r') as file:
    data = file.read()

# Use regular expressions to extract relevant data
pattern = r'(\d+)\s+([MF])\s+([\w\s]+)\s+([\d\s]+)\s+([\w\s]+)\s+([\d\s]+)\s+([\w\s]+)\s+([\d\s]+)\s+([\w\s]+)\s+([\d\s]+)\s+([\w\s]+)\s+([\d\s]+)\s+([\w\s]+)\s+([\d\s]+)\s+([\w\s]+)\s+([\w\s]+)\s+([\w\s]+)'
matches = re.findall(pattern, data)

# Define column names
columns = ['Roll No.', 'Gender', 'Name', 'Subject 1', 'Grade 1', 'Subject 2', 'Grade 2',
           'Subject 3', 'Grade 3', 'Subject 4', 'Grade 4', 'Subject 5', 'Grade 5',
           'Subject 6', 'Grade 6', 'Result', 'Compart Subject']

# Create a list of dictionaries to store the extracted data
data_list = []
for match in matches:
    data_list.append({
        'Roll No.': match[0],
        'Gender': match[1],
        'Name': match[2],
        'Subject 1': match[3],
        'Grade 1': match[4],
        'Subject 2': match[5],
        'Grade 2': match[6],
        'Subject 3': match[7],
        'Grade 3': match[8],
        'Subject 4': match[9],
        'Grade 4': match[10],
        'Subject 5': match[11],
        'Grade 5': match[12],
        'Subject 6': match[13],
        'Grade 6': match[14],
        'Result': match[15],
        'Compart Subject': match[16]
    })

# Create a pandas DataFrame from the list of dictionaries
df = pd.DataFrame(data_list, columns=columns)

# Drop any extra spaces in the 'Subject' and 'Grades' columns
df['Subject 1'] = df['Subject 1'].str.strip()
df['Subject 2'] = df['Subject 2'].str.strip()
df['Subject 3'] = df['Subject 3'].str.strip()
df['Subject 4'] = df['Subject 4'].str.strip()
df['Subject 5'] = df['Subject 5'].str.strip()
df['Subject 6'] = df['Subject 6'].str.strip()
df['Grade 1'] = df['Grade 1'].str.strip()
df['Grade 2'] = df['Grade 2'].str.strip()
df['Grade 3'] = df['Grade 3'].str.strip()
df['Grade 4'] = df['Grade 4'].str.strip()
df['Grade 5'] = df['Grade 5'].str.strip()
df['Grade 6'] = df['Grade 6'].str.strip()

# Save the DataFrame to an Excel file
df.to_excel('output.xlsx', index=False)
