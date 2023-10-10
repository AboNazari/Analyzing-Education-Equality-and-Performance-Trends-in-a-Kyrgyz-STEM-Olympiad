import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 1. Load and preprocess the data
df = pd.read_excel('allClasses.xlsx')

# Ensure there are no NaNs in necessary columns
df = df.dropna(subset=['schoolType', 'class', 'scores'])

# Convert 'scores' column to numeric, setting errors='coerce' to turn problematic strings into NaN
df['scores'] = pd.to_numeric(df['scores'], errors='coerce')

# Filter out rows with NaN scores
df = df.dropna(subset=['scores'])

# Rank the scores, and then filter the dataframe for only the top 30%
df['rank'] = df.groupby(['class', 'schoolType'])['scores'].rank(pct=True, ascending=False)
df_top_10 = df[df['rank'] <= 0.10]

# 2. Calculate the score distribution for schoolType for the top 30% for each class
grouped = df_top_10.groupby(['class', 'schoolType']).size().reset_index(name='count')

# 3. Plot the distribution for each class
classes = df['class'].unique()

for class_name in classes:
    subset = grouped[grouped['class'] == class_name]
    plt.figure(figsize=(12, 6))
    sns.barplot(x='schoolType', y='count', data=subset)
    plt.title(f"Score Distribution in Top 10% for School Type in Class: {class_name}")
    plt.ylabel('Number of Students')
    plt.tight_layout()
    plt.savefig(f"Top_10_Score_Distribution_{class_name}.png")
    plt.show()

# 4. Plot the overall distribution for all classes
overall_grouped = df_top_10.groupby(['schoolType']).size().reset_index(name='count')
plt.figure(figsize=(12, 6))
sns.barplot(x='schoolType', y='count', data=overall_grouped)
plt.title("Score Distribution in Top 10% for School Type (All Classes)")
plt.ylabel('Number of Students')
plt.tight_layout()
plt.savefig("Top_10_Score_Distribution_All_Classes.png")
plt.show()

# 5. Save the result to an Excel file
grouped.to_excel("Top_10_Score_Distribution.xlsx")
