import pandas as pd
import os

# -------------------------------
# 1. Load raw data
# -------------------------------

input_file = "data/sales_data.csv"

df = pd.read_csv(input_file)

print("Raw data loaded successfully!")
print("Original shape:", df.shape)


# -------------------------------
# 2. Handle missing values
# -------------------------------

df["Age"] = df["Age"].fillna(df["Age"].median())
df["Quantity"] = df["Quantity"].fillna(df["Quantity"].median())


# -------------------------------
# 3. Remove duplicate records
# -------------------------------

df = df.drop_duplicates()


# -------------------------------
# 4. Standardize city names
# -------------------------------

df["City"] = df["City"].str.strip().str.title()


# -------------------------------
# 5. Fix data types
# -------------------------------

df["Age"] = df["Age"].astype(int)
df["Quantity"] = df["Quantity"].astype(int)
df["Sales_Date"] = pd.to_datetime(df["Sales_Date"])


# -------------------------------
# 6. Calculate total sales
# -------------------------------

df["Total_Sales"] = df["Quantity"] * df["Unit_Price"]


# -------------------------------
# 7. Save cleaned dataset
# -------------------------------

output_file = "data/cleaned_sales_data.csv"

df.to_csv(output_file, index=False)

print("Cleaned data saved successfully!")
print("Final shape:", df.shape)
print("Total Revenue:", df["Total_Sales"].sum())
