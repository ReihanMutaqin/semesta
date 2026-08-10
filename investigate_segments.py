import sys
sys.path.append(r'C:\Users\reyha\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\site-packages')
import pandas as pd
import numpy as np

excel_path = r'C:\PROJEK\Data Semesta\export.xlsx'
df = pd.read_excel(excel_path)

print("=== SEGMENT INVESTIGATION ===")

# Check Modoroso
df['is_modoroso'] = df['Description'].astype(str).str.contains('Modoroso', case=False, na=False)
print("Modoroso count in Description:", df['is_modoroso'].sum())

# Check PDAHSI / PDA HSI / PDA
df['is_pdahsi_sc'] = df['SC Order No/Track ID/CSRM No'].astype(str).str.contains('PDA', case=False, na=False)
df['is_pdahsi_og'] = df['Owner Group'].astype(str).str.contains('PMDA', case=False, na=False)
print("PDAHSI count in SC Order No:", df['is_pdahsi_sc'].sum())
print("PMDA count in Owner Group:", df['is_pdahsi_og'].sum())

# Check Indihome
df['is_indihome_prod'] = df['Product Name'].astype(str).str.contains('INDIHOME', case=False, na=False)
df['is_indihome_pt'] = df['Product Type'] == 'COMMON'
print("INDIHOME in Product Name:", df['is_indihome_prod'].sum())
print("Product Type == COMMON:", df['is_indihome_pt'].sum())

print("\n--- Cross Tabulation of Segments ---")
df['Segment_Guess'] = 'OTHER'
df.loc[df['is_modoroso'], 'Segment_Guess'] = 'Modoroso'
df.loc[df['is_pdahsi_sc'], 'Segment_Guess'] = 'PDA HSI'
df.loc[df['is_indihome_prod'], 'Segment_Guess'] = 'IndiHome'

print(df['Segment_Guess'].value_counts(dropna=False))

print("\nRemaining OTHER sample rows:")
other_df = df[df['Segment_Guess'] == 'OTHER']
print(other_df[['Description', 'SC Order No/Track ID/CSRM No', 'Owner Group', 'Product Name', 'Product Type', 'CRM Order Type']].head(15))
