import pandas as pd

# Your dataframe
data = {
    'Roll': [22614342, 22614343, 22614344, 22614345, 22614346, 22614347],
    'Gender': ['M', 'F', 'F', 'M', 'M', 'M'],
    'Name': ['ABHIMANYU KUMAR', 'AKANKSHA PRIYA', 'AKANKSHI KUMARI', 'AKEE KUMAR', 'AMAN KUMAR', 'AMISH RANJAN'],
    '302': [80, 76, '', '', 74, '']
}

df = pd.DataFrame(data)

# Create empty dataframes to store rows
df1 = pd.DataFrame(columns=df.columns)
df2 = pd.DataFrame(columns=df.columns)

# Iterate through rows
for index, row in df.iterrows():
    if row['302'] != '':
        df1 = pd.concat([df1, pd.DataFrame([row])], ignore_index=True)
    else:
        df2 = pd.concat([df2, pd.DataFrame([row])], ignore_index=True)

# Reset index for the new dataframes
# df1.reset_index(drop=True, inplace=True)
# df2.reset_index(drop=True, inplace=True)

print("New df1:")
print(df1)

print("\nNew df2:")
print(df2)
