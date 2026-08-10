import sys
sys.path.append(r'C:\Users\reyha\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\site-packages')
import pandas as pd

excel_path = r'C:\PROJEK\Data Semesta\export.xlsx'
df = pd.read_excel(excel_path)

print("Total Rows:", len(df))
print("\nDescription non-null count:", df['Description'].notna().sum())
print("Sample Descriptions (first 20):")
print(df['Description'].head(20).to_list())

print("\nValue counts of Description (top 15):")
print(df['Description'].value_counts(dropna=False).head(15))

print("\nValue counts of SC Order No/Track ID/CSRM No prefix/sample:")
print(df['SC Order No/Track ID/CSRM No'].astype(str).str[:10].value_counts().head(10))

print("\nWO Class:")
print(df['WO Class'].value_counts(dropna=False))

print("\nProduct Name top 15:")
print(df['Product Name'].value_counts(dropna=False).head(15))

print("\nProduct Type:")
print(df['Product Type'].value_counts(dropna=False))

print("\nOwner Group top 15:")
print(df['Owner Group'].value_counts(dropna=False).head(15))

print("\nCRM Order Type:")
print(df['CRM Order Type'].value_counts(dropna=False))

print("\nStatus:")
print(df['Status'].value_counts(dropna=False))
