# Import libraries
import pandas as pd
from datetime import datetime
import os
from dateutil.relativedelta import relativedelta

# Read both Excel files
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

emp_df = pd.read_excel(os.path.join(BASE_DIR, "data", "employee_data.xlsx"))
email_df = pd.read_excel(os.path.join(BASE_DIR, "data", "employee_emails.xlsx"))


# Merge using EmpID
data = emp_df.merge(email_df, on='EmpID', how='left')
print("Data merged successfully!")
data.head()

# Calculate Annual Salary
data['Annual Salary'] = data['Monthly Salary'] * 12

# Calculate Tenure in months based on StartDate
today = datetime.today()

def months_between(start, end):
    if pd.isna(start):
        return None
    rd = relativedelta(end, start)
    return rd.years * 12 + rd.months

data['Tenure (months)'] = data['StartDate'].apply(lambda x: months_between(x, today))

# Create Tenure Group
def tenure_group(months):
    if months is None:
        return None
    if months <= 6:
        return '0-6 months'
    elif months <= 24:
        return '1-2 years'
    elif months <= 36:
        return '2-3 years'
    elif months <= 48:
        return '3-4 years'
    elif months <= 60:
        return '4-5 years'
    else:
        return '>5 years'

data['Tenure Group'] = data['Tenure (months)'].apply(tenure_group)


import re

# Create output folder path
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# Create output folder if it doesn't exist
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------
# Export master file
# -------------------------------
master_file_path = os.path.join(OUTPUT_DIR, "Employee Master Data.xlsx")
data.to_excel(master_file_path, index=False)
print(f"Employee Master Data.xlsx created at: {OUTPUT_DIR}")

# -------------------------------
# Split data by DepartmentType
# -------------------------------
for dept, df in data.groupby('DepartmentType'):
    # Replace special characters in department name
    safe_name = re.sub(r'[\\/*?:"<>|]', '_', str(dept)).strip()
    file_name = f"Employee Master Data - {safe_name}.xlsx"
    file_path = os.path.join(OUTPUT_DIR, file_name)

    df.to_excel(file_path, index=False)
    print(f"Created file: {file_path}")


