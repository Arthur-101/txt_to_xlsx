import pandas as pd

# Create a sample dataframe
data = {'Value': [10, 20, 15, 25, 30]}
df = pd.DataFrame(data)

# Add a new column with the previous row's value
df['Previous_Value'] = df['Value'].shift(1)

print(df)
