import pandas as pd

# Load the business dataset
df = pd.read_csv("data/business_sales_data.csv")

print("===== DATA CLEANING STARTED =====")

# 1. Check basic information
print("\nDataset Shape:")
print(df.shape)

# 2. Check missing values
print("\nMissing Values:")
print(df.isnull().sum())

# 3. Check duplicate records
print("\nDuplicate Records:")
print(df.duplicated().sum())

# 4. Convert Date column to datetime
df["Date"] = pd.to_datetime(df["Date"])

# 5. Remove duplicate records
df = df.drop_duplicates()

# 6. Remove rows with missing values
df = df.dropna()

# 7. Sort data by date
df = df.sort_values("Date")

# 8. Save cleaned dataset
df.to_csv("data/cleaned_business_sales.csv", index=False)

print("\n===== DATA CLEANING COMPLETED =====")
print("Cleaned Records:", len(df))
print("Cleaned File: data/cleaned_business_sales.csv")