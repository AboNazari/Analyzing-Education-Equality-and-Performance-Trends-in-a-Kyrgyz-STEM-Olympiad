import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 1. Load and preprocess the data
df = pd.read_excel('allClasses.xlsx')

# Ensure there are no NaNs in 'schoolType', 'region', and 'class' columns
df = df.dropna(subset=['schoolType', 'region', 'class'])

# 2. Calculate the school type distribution for each class in each region
grouped = df.groupby(['class', 'region', 'schoolType']
                     ).size().reset_index(name='count')

# Save the grouped results to an Excel file
grouped.to_excel('grouped_results.xlsx', index=False)

# 3. Plot the distribution for each class and save the graphs
classes = df['class'].unique()

for idx, class_name in enumerate(classes, 1):
    subset = grouped[grouped['class'] == class_name]
    plt.figure(figsize=(12, 6))
    sns.barplot(x='region', y='count', hue='schoolType', data=subset)
    plt.title(
        f"School Type Distribution in Each Region for Class: {class_name}")
    plt.ylabel('Number of Schools')
    plt.tight_layout()
    plt.savefig(f'class_{idx}_distribution.png')
    plt.show()

# 4. Plot the overall distribution for all classes and save the graph
overall_grouped = df.groupby(
    ['region', 'schoolType']).size().reset_index(name='count')
plt.figure(figsize=(12, 6))
sns.barplot(x='region', y='count', hue='schoolType', data=overall_grouped)
plt.title("Overall School Type Distribution in Each Region")
plt.ylabel('Number of Schools')
plt.tight_layout()
plt.savefig('overall_distribution.png')
plt.show()
