import pandas as pd
import numpy as np

# Read the excel file 'everything.xlsx'
everything_df = pd.read_excel('everything.xlsx')

# Create a mapping of 'name' to 'region' where region is 'Bishkek'
name_region_mapping = everything_df[everything_df['region'] == 'Bishkek'].set_index('name')['region'].to_dict()

# Read the excel file 'allClasses.xlsx'
allClasses_df = pd.read_excel('allClasses.xlsx')

# Create a new column 'new_region' based on the mapping. It will be NaN for 'names' not found in the mapping
allClasses_df['new_region'] = allClasses_df['name'].map(name_region_mapping)

# Create a mask where region is not 'N/A'
not_na_mask = allClasses_df['region'].ne('N/A')

# Update the 'region' column only where 'new_region' is not NaN and original 'region' is not 'N/A'
allClasses_df.loc[(~allClasses_df['new_region'].isna()) & not_na_mask, 'region'] = allClasses_df['new_region']

# Drop the 'new_region' column
allClasses_df.drop('new_region', axis=1, inplace=True)

# Fill all NaN fields with 'N/A'
allClasses_df.fillna('N/A', inplace=True)

# Write the updated DataFrame to an Excel file
allClasses_df.to_excel('updated_allClasses.xlsx', index=False)
