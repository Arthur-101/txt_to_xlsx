import csv

# Open a CSV file for writing
with open('data.csv', 'w', newline='') as file:
    # Create a CSV writer object
    writer = csv.writer(file)

    # Write the header row
    writer.writerow(['Column1', 'Column2', 'Column3'])

    # Write data rows
    writer.writerow(['Value1', 'Value2', 'Value3'])
    writer.writerow(['Value4', 'Value5', 'Value6'])

# Open a CSV file for reading
with open('data.csv', 'r') as file:
    # Create a CSV reader object
    reader = csv.reader(file)

    # Read the header row
    header = next(reader)

    # Iterate over each row in the CSV file
    for row in reader:
        # Accessing data in each column of the current row
        column1 = row[0]
        column2 = row[1]

# Open a CSV file for reading
with open('data.csv', 'r') as file:
    # Create a CSV dictionary reader object
    reader = csv.DictReader(file)

    # Iterate over each row in the CSV file
    for row in reader:
        # Accessing data using column names as dictionary keys
        column1 = row['Column1']
        column2 = row['Column2']

# Open a CSV file for writing
with open('data.csv', 'w', newline='') as file:
    # Define field names for the CSV file
    fieldnames = ['Column1', 'Column2', 'Column3']

    # Create a CSV dictionary writer object
    writer = csv.DictWriter(file, fieldnames=fieldnames)

    # Write the header row
    writer.writeheader()

    # Write data rows
    writer.writerow({'Column1': 'Value1', 'Column2': 'Value2', 'Column3': 'Value3'})
    writer.writerow({'Column1': 'Value4', 'Column2': 'Value5', 'Column3': 'Value6'})

# Register a custom CSV dialect
csv.register_dialect('my_dialect', delimiter=';', quoting=csv.QUOTE_ALL)

# Open a CSV file using the custom dialect
with open('data.csv', 'r') as file:
    reader = csv.reader(file, dialect='my_dialect')
    
# Open a CSV file for reading
with open('data.csv', 'r') as file:
    reader = csv.reader(file)
    header = next(reader)
    print(header)  # Output: ['Column1', 'Column2', 'Column3']
    
