import pandas as pd
import matplotlib.pyplot as plt

# Read the Excel data into a DataFrame
data = pd.read_excel('allClasses.xlsx')

# Convert 'scores' column to numeric, coerce errors to NaN
data['scores'] = pd.to_numeric(data['scores'], errors='coerce')

# Drop rows with NaN 'scores'
data.dropna(subset=['scores'], inplace=True)

# Sort the data by scores in descending order
data.sort_values(by='scores', ascending=False, inplace=True)

# Initialize an empty list to store the terminal results
result_rows = []

# Create an empty DataFrame to aggregate data across all classes
aggregate_data = pd.DataFrame()

# Iterate over each class
for class_name, class_data in data.groupby('class'):
    # Calculate the number of rows corresponding to the top 30% scores in the current class
    top_30_percent = int(len(class_data) * 0.3)

    # Get the top 30% scores for the current class
    top_scores = class_data.head(top_30_percent)
    
    # Add this data to the aggregate data frame
    aggregate_data = pd.concat([aggregate_data, top_scores])

    # Calculate the distribution of regions in the top scores for the current class
    region_distribution = top_scores['region'].value_counts()

    # Calculate the percentage of each region in the top scores for the current class
    region_percentage = region_distribution / region_distribution.sum() * 100

    # Save distribution data
    for region, count in region_distribution.items():
        percentage = region_percentage[region]
        result_rows.append({'Class': class_name, 'Region': region, 'Count': count, 'Percentage': percentage})

    # Plot the percentage distribution of regions in a pie chart for the current class
    plt.pie(region_percentage, labels=region_percentage.index, autopct='%1.1f%%')
    plt.title(f'Distribution of Regions in Top 30% Scores - Class {class_name}')
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    plt.savefig(f'Class_{class_name}_pie_chart_30.png')
    plt.close()

# Calculate overall region distribution across all classes
overall_region_distribution = aggregate_data['region'].value_counts()
overall_region_percentage = overall_region_distribution / overall_region_distribution.sum() * 100

# Plot the overall region distribution
plt.pie(overall_region_percentage, labels=overall_region_percentage.index, autopct='%1.1f%%')
plt.title('Distribution of Regions in Top 30% Scores - All Classes Combined')
plt.axis('equal')
plt.savefig('All_Classes_pie_chart_30.png')
plt.close()

# Create a DataFrame from the result_rows list
result_table = pd.DataFrame(result_rows, columns=['Class', 'Region', 'Count', 'Percentage'])

# Save the result_table DataFrame to an Excel file
result_table.to_excel('result_table_30.xlsx', index=False)
