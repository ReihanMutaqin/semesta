import sys
sys.path.append(r'C:\Users\reyha\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\site-packages')
import pandas as pd

excel_path = r'C:\PROJEK\Data Semesta\export.xlsx'
df = pd.read_excel(excel_path)

sc = df['SC Order No/Track ID/CSRM No'].astype(str)
print("=== Top 30 SC Order No Prefixes (first 4 chars) ===")
print(sc.str[:4].value_counts())

print("\n=== Top 30 SC Order No Prefixes (first 6 chars) ===")
print(sc.str[:6].value_counts().head(30))

print("\n=== Owner Group Value Counts ===")
print(df['Owner Group'].value_counts())
