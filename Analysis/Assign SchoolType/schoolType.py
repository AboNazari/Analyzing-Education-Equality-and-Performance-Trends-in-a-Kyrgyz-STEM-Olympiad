import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.tree import DecisionTreeClassifier
from sklearn.feature_extraction.text import CountVectorizer

# Load data from extractedSchools.xlsx
df = pd.read_excel('extractedSchools.xlsx')

# Remove rows where 'school' or 'Type' columns have NaN
df = df.dropna(subset=['school', 'Type'])

# Convert the 'school' column to string
df['school'] = df['school'].astype(str)

# Convert the 'school' column to a matrix of token counts
vectorizer = CountVectorizer()
X = vectorizer.fit_transform(df['school'])

# Convert 'Type' column to target variable
y = df['Type']

# Split the data into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Train a decision tree classifier
clf = DecisionTreeClassifier()
clf.fit(X_train, y_train)

# Test the classifier's accuracy
accuracy = clf.score(X_test, y_test)
print(f"Accuracy: {accuracy:.2f}")

# Load data from allClasses.xlsx
all_classes_df = pd.read_excel('allClasses.xlsx')

# Handle NaNs in 'school' column by replacing them with a placeholder
all_classes_df['school'] = all_classes_df['school'].fillna("unknown_school")

# Convert the 'school' column in allClasses.xlsx to string
all_classes_df['school'] = all_classes_df['school'].astype(str)

# Transform 'school' column from the new dataframe using the same vectorizer
X_new = vectorizer.transform(all_classes_df['school'])

# Predict school type for all rows in allClasses dataframe
all_classes_df['schoolType'] = clf.predict(X_new)

# Save the results back to Excel, if needed
all_classes_df.to_excel('allClasses_updated.xlsx', index=False)
