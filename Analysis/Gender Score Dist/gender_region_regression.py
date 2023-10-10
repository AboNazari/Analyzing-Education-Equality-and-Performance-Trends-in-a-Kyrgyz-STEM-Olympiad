import pandas as pd
import matplotlib.pyplot as plt

# Read the Excel data into a DataFrame
data = pd.read_excel('allClasses.xlsx')

# Convert 'scores' column to numeric, coerce errors to NaN
data['scores'] = pd.to_numeric(data['scores'], errors='coerce')

# Drop rows with NaN 'scores'
data.dropna(subset=['scores'], inplace=True)

# Initialize a dictionary to store the region and gender distribution for each class
class_distribution = {}

# Initialize an empty list to store the terminal results
result_rows = []

# Iterate over each class
for class_name, class_data in data.groupby('class'):
    # Sort the data for the current class by scores in descending order
    class_data = class_data.sort_values(by='scores', ascending=False)

    # Calculate the number of rows corresponding to the top 30% scores in the current class
    top_30_percent = int(len(class_data) * 0.3)

    # Get the top 30% scores for the current class
    top_scores = class_data.head(top_30_percent)

    # Calculate the gender and region distribution in the top scores for the current class
    region_gender_distribution = top_scores.groupby(['region', 'gender']).size().unstack(fill_value=0)

    # Store the region and gender distribution in the dictionary
    class_distribution[class_name] = region_gender_distribution

    # Append the results to the result_rows list
    for region in region_gender_distribution.index:
        count_male = region_gender_distribution.loc[region, 'Male']
        count_female = region_gender_distribution.loc[region, 'Female']
        result_rows.append({'Class': class_name,
                            'Region': region,
                            'Gender': 'Male',
                            'Count': count_male})
        result_rows.append({'Class': class_name,
                            'Region': region,
                            'Gender': 'Female',
                            'Count': count_female})

# Create a DataFrame from the result_rows list
result_table = pd.DataFrame(result_rows, columns=['Class', 'Region', 'Gender', 'Count'])

# Save the result_table DataFrame to an Excel file
result_table.to_excel('result_table.xlsx', index=False)

# Plot the score distribution in each region for each gender within the top 30%
for class_name, region_gender_distribution in class_distribution.items():
    fig, ax = plt.subplots()
    region_labels = region_gender_distribution.index
    num_regions = len(region_labels)
    bar_width = 0.35
    index = range(num_regions)

    male_scores = region_gender_distribution['Male']
    female_scores = region_gender_distribution['Female']

    ax.bar(index, male_scores, bar_width, color='tab:blue', label='Male')
    ax.bar(index, female_scores, bar_width, color='tab:pink',
           label='Female', bottom=male_scores)

    ax.set_xlabel('Region')
    ax.set_ylabel('Count')
    ax.set_title(f'Score Distribution by Region and Gender - Class {class_name}')
    ax.set_xticks(index)
    ax.set_xticklabels(region_labels)
    ax.legend()
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f'Class_{class_name}_score_distribution.png')
    plt.close()

# For all classes combined
combined_distribution = data.groupby(['region', 'gender']).size().unstack(fill_value=0)
fig, ax = plt.subplots()
region_labels = combined_distribution.index
num_regions = len(region_labels)
bar_width = 0.35
index = range(num_regions)

male_scores = combined_distribution['Male']
female_scores = combined_distribution['Female']

ax.bar(index, male_scores, bar_width, color='tab:blue', label='Male')
ax.bar(index, female_scores, bar_width, color='tab:pink', label='Female', bottom=male_scores)

ax.set_xlabel('Region')
ax.set_ylabel('Count')
ax.set_title('Score Distribution by Region and Gender - All Classes Combined')
ax.set_xticks(index)
ax.set_xticklabels(region_labels)
ax.legend()
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('All_Classes_score_distribution.png')
plt.close()
