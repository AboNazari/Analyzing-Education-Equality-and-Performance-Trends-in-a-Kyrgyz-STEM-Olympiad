import pandas as pd
import matplotlib.pyplot as plt

# Read the Excel data into a DataFrame
data = pd.read_excel('allClasses.xlsx')

# Convert 'scores' column to numeric, coerce errors to NaN
data['scores'] = pd.to_numeric(data['scores'], errors='coerce')

# Drop rows with NaN 'scores'
data.dropna(subset=['scores'], inplace=True)

# Only consider rows with 'gender' as 'Male' or 'Female'
data = data[data['gender'].isin(['Male', 'Female'])]

# Initialize an empty list to store the terminal results
result_rows = []

# For all classes combined
aggregated_gender_distribution = data.groupby('gender').size()
aggregated_gender_distribution_percentage = aggregated_gender_distribution / aggregated_gender_distribution.sum() * 100

# Plot the aggregated gender distribution
aggregated_gender_distribution_percentage.plot(kind='bar', color=['blue', 'pink'])
plt.title('Distribution of Genders Across All Classes and Regions')
plt.ylabel('Percentage')
plt.xticks(rotation=0)
plt.savefig('All_Classes_And_Regions_gender_distribution.png')
plt.close()

# Iterate over each class and region
for (class_name, region_name), group_data in data.groupby(['class', 'region']):
    # Sort the data for the current group by scores in descending order
    group_data = group_data.sort_values(by='scores', ascending=False)

    # Calculate the number of rows corresponding to the top 30% scores in the current group
    top_30_percent = int(len(group_data) * 0.3)

    # Get the top 30% scores for the current group
    top_scores = group_data.head(top_30_percent)

    # Calculate the distribution of genders in the top scores for the current group
    gender_distribution = top_scores['gender'].value_counts()

    # Calculate the percentage of each gender in the top scores for the current group
    gender_percentage = gender_distribution / gender_distribution.sum() * 100

    # Append the results to the result_rows list
    for gender, count in gender_distribution.items():
        percentage = gender_percentage[gender]
        result_rows.append({'Class': class_name,
                            'Region': region_name,
                            'Gender': gender,
                            'Count': count,
                            'Percentage': percentage})

# Create a DataFrame from the result_rows list
result_table = pd.DataFrame(result_rows, columns=['Class', 'Region', 'Gender', 'Count', 'Percentage'])

# Save the result_table DataFrame to an Excel file
result_table.to_excel('gender_result_table.xlsx', index=False)
