import sys
sys.path.append(r'C:\Users\reyha\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\site-packages')
import pandas as pd
import numpy as np

excel_path = r'C:\PROJEK\Data Semesta\export.xlsx'
df = pd.read_excel(excel_path)

print('=== DATASET OVERVIEW ===')
print('Total Rows:', len(df))

print('\n--- Search Keywords Across Columns ---')
for kw in ['modoroso', 'pdahsi', 'pdah', 'pmda', 'indihome', 'hsi', 'tif', 'enterprise', 'common']:
    matches = {}
    for col in df.columns:
        cnt = df[col].astype(str).str.contains(kw, case=False, na=False).sum()
        if cnt > 0:
            matches[col] = cnt
    print(f'Keyword "{kw}": {matches}')

print('\n--- Owner Group Unique Values ---')
print(df['Owner Group'].value_counts(dropna=False))

print('\n--- Product Type Unique Values ---')
print(df['Product Type'].value_counts(dropna=False))

print('\n--- CRM Order Type Unique Values ---')
print(df['CRM Order Type'].value_counts(dropna=False))

print('\n--- Status Unique Values ---')
print(df['Status'].value_counts(dropna=False))

print('\n--- Sample Descriptions ---')
print(df['Description'].dropna().head(10))
