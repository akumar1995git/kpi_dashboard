import pandas as pd
import matplotlib.pyplot as plt

# Replace with your actual Excel file name or path
file_path = 'your_uploaded_file.xlsx'

# Read in the data from the first sheet
df = pd.read_excel(file_path)

# Print all column names for reference
print("Column names in the uploaded Excel file:")
for col in df.columns:
    print(col)

# ---- Example Recommended Charts ----
# Update column names below to match those printed above

# 1. Bar chart: Department vs. Employee Count
if 'Department' in df.columns:
    dept_count = df['Department'].value_counts()
    plt.figure(figsize=(8,5))
    dept_count.plot(kind='bar')
    plt.title("Employees Per Department")
    plt.xlabel("Department")
    plt.ylabel("Count")
    plt.tight_layout()
    plt.show()

# 2. Line chart: Monthly Burnout Score Trend
if {'Month', 'BurnoutScore'}.issubset(df.columns):
    burnout_monthly = df.groupby('Month')['BurnoutScore'].mean()
    plt.figure(figsize=(8,5))
    burnout_monthly.plot(kind='line', marker='o')
    plt.title("Average Burnout Score by Month")
    plt.xlabel("Month")
    plt.ylabel("Burnout Score")
    plt.tight_layout()
    plt.show()

# 3. Boxplot: Productivity by Department
if {'Department', 'ProductivityScore'}.issubset(df.columns):
    plt.figure(figsize=(8,5))
    df.boxplot(column='ProductivityScore', by='Department')
    plt.title("Productivity Score Distribution by Department")
    plt.suptitle('')  # Removes default super title
    plt.xlabel("Department")
    plt.ylabel("Productivity Score")
    plt.tight_layout()
    plt.show()

# 4. Scatter plot: Burnout vs. Productivity
if {'BurnoutScore', 'ProductivityScore'}.issubset(df.columns):
    plt.figure(figsize=(8,5))
    plt.scatter(df['BurnoutScore'], df['ProductivityScore'])
    plt.title("Burnout Score vs Productivity Score")
    plt.xlabel("Burnout Score")
    plt.ylabel("Productivity Score")
    plt.tight_layout()
    plt.show()

# 5. Histogram: Meeting/Communication Hours
for col in df.columns:
    if 'Meeting' in col or 'Communication' in col:
        plt.figure(figsize=(8,5))
        df[col].plot(kind='hist', bins=20)
        plt.title(f"Distribution of {col}")
        plt.xlabel(col)
        plt.ylabel("Frequency")
        plt.tight_layout()
        plt.show()

# 6. Pie Chart: Department Share
if 'Department' in df.columns:
    plt.figure(figsize=(8,5))
    dept_count.plot(kind='pie', autopct='%1.1f%%')
    plt.title("Department Share")
    plt.ylabel("")
    plt.tight_layout()
    plt.show()

# --- Add/Adjust chart code as needed for other analytics KPIs ---
