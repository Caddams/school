import pandas as pd
import os
import numpy as np
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.neighbors import KNeighborsClassifier
from sklearn.metrics import accuracy_score, classification_report

# Load the CSV file into a DataFrame
directory = r'C:\Users\RedMa\OneDrive\Desktop\Skewl'
file_name = 'CHI_Crime.csv'
file_path = os.path.join(directory, file_name)

# Read the CSV file with error handling
try:
    df = pd.read_csv(file_path, on_bad_lines='warn', engine='python')
except pd.errors.ParserError as e:
    print(f"Error parsing CSV file: {e}")

# Drop rows with missing values
df.dropna(inplace=True)

# Convert categorical variables to numerical ones
df['Primary Type'] = df['Primary Type'].astype('category').cat.codes

# Select relevant features
features = ['Latitude', 'Longitude', 'Primary Type']
df = df[features]

# Check for missing values and data types
print(df.isnull().sum())
print(df.dtypes)
print(f'Total samples: {len(df)}')

X = df[['Latitude', 'Longitude']]
y = df['Primary Type']

# Adjust the test_size parameter if the dataset is small
if len(df) > 1:
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.1, random_state=42)
else:
    raise ValueError("Dataset is too small to split.")

# Initialize the KNN model
knn = KNeighborsClassifier(n_neighbors=5)

# Train the model
knn.fit(X_train, y_train)

# Make predictions
y_pred = knn.predict(X_test)

# Calculate accuracy
accuracy = accuracy_score(y_test, y_pred)
print(f'Accuracy: {accuracy * 100:.2f}%')

# Print classification report
print(classification_report(y_test, y_pred))

# Example new data point
new_data = np.array([[41.8781, -87.6298]])  # Latitude and Longitude of a new location

# Predict the crime type
predicted_crime = knn.predict(new_data)
print(f'Predicted Crime Type: {predicted_crime[0]}')

# Optionally, use cross-validation for small datasets
scores = cross_val_score(knn, X, y, cv=5)
print(f'Cross-validation scores: {scores}')
print(f'Mean cross-validation score: {scores.mean()}')
