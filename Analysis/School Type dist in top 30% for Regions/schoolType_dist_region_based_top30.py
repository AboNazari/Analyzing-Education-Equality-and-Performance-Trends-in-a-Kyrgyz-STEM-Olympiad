import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 1. Load and preprocess the data
df = pd.read_excel('allClasses.xlsx')

# Ensure there are no NaNs in necessary columns
df = df.dropna(subset=['schoolType', 'region', 'scores'])

# Convert 'scores' column to numeric
df['scores'] = pd.to_numeric(df['scores'], errors='coerce')

# Filter out rows with NaN scores
df = df.dropna(subset=['scores'])

# Rank the scores, and then filter the dataframe for only the top 30%
df['rank'] = df.groupby(['region', 'schoolType'])['scores'].rank(pct=True, ascending=False)
df_top_30 = df[df['rank'] <= 0.30]

# 2. Calculate the score distribution for schoolType for the top 30% for each region
grouped = df_top_30.groupby(['region', 'schoolType']).size().reset_index(name='count')

# 3. Plot the distribution for each region
regions = df['region'].unique()

for region_name in regions:
    subset = grouped[grouped['region'] == region_name]
    plt.figure(figsize=(12, 6))
    sns.barplot(x='schoolType', y='count', data=subset)
    plt.title(f"Score Distribution in Top 30% for School Type in Region: {region_name}")
    plt.ylabel('Number of Students')
    plt.tight_layout()
    plt.savefig(f"Top_30_Score_Distribution_Region_{region_name}.png")
    plt.show()

# 4. Plot the overall distribution for all regions
overall_grouped = df_top_30.groupby(['schoolType']).size().reset_index(name='count')
plt.figure(figsize=(12, 6))
sns.barplot(x='schoolType', y='count', data=overall_grouped)
plt.title("Score Distribution in Top 30% for School Type (All Regions)")
plt.ylabel('Number of Students')
plt.tight_layout()
plt.savefig("Top_30_Score_Distribution_All_Regions.png")
plt.show()

# 5. Save the result to an Excel file
grouped.to_excel("Top_30_Score_Distribution_Regions.xlsx")
