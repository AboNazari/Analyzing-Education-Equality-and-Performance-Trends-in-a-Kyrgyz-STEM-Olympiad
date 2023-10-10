import pandas as pd

# Read the excel file 'updated_allClasses.xlsx'
df = pd.read_excel('updated_allClasses.xlsx')

# Convert 'city' and 'region' columns to lowercase and strip spaces
df['city'] = df['city'].str.strip()
df['region'] = df['region'].str.strip()

# Update 'region' to 'Bishkek' for rows where 'city' is 'Bishkek' and 'region' is not 'Bishkek'
df.loc[(df['city'] == 'Bishkek') & (
    df['region'] != 'Bishkek'), 'region'] = 'Bishkek'
# Fill all NaN fields with 'N/A'
df.fillna('N/A', inplace=True)
# Write the updated DataFrame to an Excel file
df.to_excel('updated_allClasses.xlsx', index=False)
