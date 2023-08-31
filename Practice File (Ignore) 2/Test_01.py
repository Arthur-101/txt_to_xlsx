import pandas as pd

# Your original DataFrame
data = {
    'Roll': [22614342, 22614343, 22614368, 22614369, 22614370, 22614388, 22614389, 22614407, 22614408, 22614440, 22614441, 22614451, 22614452, 22614467, 22614468],
    'Gender': ['M', 'F', 'M', 'M', 'F', 'F', 'F', 'M', 'F', 'M', 'F', 'M', 'M', 'M', 'M'],
    'Name': ['ABHIMANYU KUMAR', 'AKANKSHA PRIYA', 'MD FARHAN ALI', 'OJAS SINHA', 'PALAK KUMARI SINHA', 'UNNATI KUMARI', 'VAISHNAVI', 'AYUSH KUMAR SINGH', 'DERSHITA GAUTAM', 'SWAPNIL KUMAR', 'VIDUSHI PRIYA', 'DURGESH KUMAR', 'GAURAV KUMAR', 'VIKASH KUMAR MANN', 'YASH RAJ']
}

df = pd.DataFrame(data)

# Sorting the DataFrame by Name
df_sorted = df.sort_values(by='Name')

# Finding the index where the first letter is closest to 'A' and 'Z'
closest_to_a = df_sorted['Name'].apply(lambda x: ord(x[0]) - ord('A'))
closest_to_z = df_sorted['Name'].apply(lambda x: ord('Z') - ord(x[0]))

# Creating a new column to identify groups based on proximity to 'A' and 'Z'
df_sorted['Group'] = closest_to_a < closest_to_z

# Creating the separate DataFrames based on the 'Group' column
grouped_dfs = df_sorted.groupby('Group')

# Printing the resulting DataFrames
for group, group_df in grouped_dfs:
    print(f"DataFrame {group + 1}:\n{group_df}\n\n")
