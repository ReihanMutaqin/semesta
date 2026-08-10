import sys
sys.path.append(r'C:\Users\reyha\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\site-packages')
import pandas as pd
import numpy as np

excel_path = r'C:\PROJEK\Data Semesta\export.xlsx'
df = pd.read_excel(excel_path)

print("=== DETAILED PREFIX AND CLASSIFICATION TEST ===")

sc_str = df['SC Order No/Track ID/CSRM No'].astype(str)
og_str = df['Owner Group'].astype(str)
prod_str = df['Product Name'].astype(str)
pt_str = df['Product Type'].astype(str)

print("MYIR- prefix count in SC Order No:", sc_str.str.startswith('MYIR').sum())
print("DGPS- / PDA prefix count in SC Order No:", (sc_str.str.contains('DGPS', case=False) | sc_str.str.contains('PDA', case=False)).sum())
print("1-D / 1- prefix count in SC Order No:", sc_str.str.startswith('1-').sum())

# Let's inspect SC Order No top prefixes (first 4 characters)
print("\nTop 20 SC Order No 4-char prefixes:")
print(sc_str.str[:4].value_counts().head(20))

# Let's inspect Product Type by SC Order No prefix
df['SC_Prefix'] = sc_str.str[:4]
print("\nProduct Type by SC_Prefix top 10:")
print(pd.crosstab(df['SC_Prefix'], df['Product Type'].fillna('MISSING'), dropna=False).head(15))

print("\nOwner Group by SC_Prefix top 10:")
print(pd.crosstab(df['SC_Prefix'], df['Owner Group'].fillna('MISSING'), dropna=False).head(15))
