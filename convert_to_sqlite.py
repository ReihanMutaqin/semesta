import sys
sys.path.append(r'C:\Users\reyha\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\site-packages')
import pandas as pd
import numpy as np
import sqlite3
import json
import os

excel_path = r'C:\PROJEK\Data Semesta\export.xlsx'
db_path = r'C:\PROJEK\Data Semesta\semesta.db'
json_summary_path = r'C:\PROJEK\Data Semesta\summary.json'

print("Step 1: Reading export.xlsx into DataFrame...")
df = pd.read_excel(excel_path)
print(f"Loaded {len(df)} rows.")

print("Step 2: Pre-processing and cleaning columns...")
df['Date Created'] = pd.to_datetime(df['Date Created'], errors='coerce')
df['Status Date'] = pd.to_datetime(df['Status Date'], errors='coerce')
df['Date Modified'] = pd.to_datetime(df['Date Modified'], errors='coerce')
df['Sched Start'] = pd.to_datetime(df['Sched Start'], errors='coerce')
df['Booking Date'] = pd.to_datetime(df['Booking Date'], errors='coerce')
df['Measurement Date'] = pd.to_datetime(df['Measurement Date'], errors='coerce')

max_date = max(df['Date Created'].max(), df['Status Date'].max())

# Standardize CRM Order Type
df['CRM_Order_Type_Raw'] = df['CRM Order Type'].fillna('UNSPECIFIED').astype(str).str.strip()
df['Tipe_Transaksi'] = df['CRM_Order_Type_Raw'].str.upper()

crm_map = {
    'NEW INSTALL': 'CREATE',
    'CREATE': 'CREATE',
    'MODIFY': 'MODIFY',
    'DISCONNECT': 'DISCONNECT',
    'SUSPEND': 'SUSPEND',
    'RESUME': 'RESUME',
    'MIGRATE': 'MIGRATE',
    'MOVE': 'MOVE'
}
df['Tipe_Transaksi_Std'] = df['Tipe_Transaksi'].map(lambda x: crm_map.get(x, x))

ps_statuses = ['COMPLETE', 'COMPWORK', 'INSTCOMP', 'DEINSTCOMP', 'ACTCOMP', 'VALCOMP']
df['Status_Clean'] = df['Status'].fillna('UNKNOWN').astype(str).str.upper().str.strip()
df['Is_PS'] = df['Status_Clean'].isin(ps_statuses).astype(int)

# PS Duration in days
df['PS_Duration_Days'] = np.where(
    (df['Is_PS'] == 1) & df['Status Date'].notna() & df['Date Created'].notna(),
    (df['Status Date'] - df['Date Created']).dt.total_seconds() / 86400.0,
    np.nan
)
df['PS_Duration_Days'] = df['PS_Duration_Days'].apply(lambda x: x if pd.notna(x) and x >= 0 else np.nan)

# Active Pending Duration in days
df['Active_Duration_Days'] = np.where(
    (df['Is_PS'] == 0) & df['Date Created'].notna(),
    (max_date - df['Date Created']).dt.total_seconds() / 86400.0,
    np.nan
)

one_month_ago = max_date - pd.Timedelta(days=30)
df['Is_PS_Last_Month'] = ((df['Is_PS'] == 1) & (df['Status Date'] >= one_month_ago)).astype(int)

# Segment Classification
def get_segment(row):
    sc = str(row['SC Order No/Track ID/CSRM No'])
    og = str(row['Owner Group'])
    pname = str(row['Product Name'])
    ptype = str(row['Product Type'])
    
    if 'DGPS' in sc or 'PDA' in sc or 'PMDA' in og:
        return 'PDA HSI'
    elif 'INDIHOME' in pname or ptype == 'COMMON':
        return 'IndiHome'
    elif 'TIF FBB' in og or sc.startswith(('MYIR', 'SC10', 'SC20', '801M', '802M', '803M', 'C001', 'C002', 'A301')):
        return 'Modoroso'
    else:
        return 'Enterprise / Lainnya'

df['Segment'] = df.apply(get_segment, axis=1)

# Format dates to string for SQLite
df['date_created_str'] = df['Date Created'].dt.strftime('%Y-%m-%d %H:%M:%S')
df['status_date_str'] = df['Status Date'].dt.strftime('%Y-%m-%d %H:%M:%S')
df['date_modified_str'] = df['Date Modified'].dt.strftime('%Y-%m-%d %H:%M:%S')
df['sched_start_str'] = df['Sched Start'].dt.strftime('%Y-%m-%d %H:%M:%S')
df['booking_date_str'] = df['Booking Date'].dt.strftime('%Y-%m-%d %H:%M:%S')
df['measurement_date_str'] = df['Measurement Date'].dt.strftime('%Y-%m-%d %H:%M:%S')

print("Step 3: Creating SQLite database with all 31 detailed columns...")
if os.path.exists(db_path):
    os.remove(db_path)

conn = sqlite3.connect(db_path)
cur = conn.cursor()

db_df = pd.DataFrame({
    'date_created': df['date_created_str'],
    'date_modified': df['date_modified_str'],
    'workorder': df['Workorder'].fillna(''),
    'sc_order_no': df['SC Order No/Track ID/CSRM No'].fillna(''),
    'oss_order_id': df['OSS Order ID'].fillna(''),
    'service_no': df['Service No.'].fillna(''),
    'description': df['Description'].fillna(''),
    'crm_order_type': df['Tipe_Transaksi_Std'],
    'owner_group': df['Owner Group'].fillna(''),
    'status': df['Status_Clean'],
    'product_name': df['Product Name'].fillna(''),
    'address': df['Address'].fillna(''),
    'witel': df['witel'].fillna(''),
    'customer_name': df['Customer Name'].fillna(''),
    'workzone': df['Workzone'].fillna(''),
    'region_site_id': df['Region/Site ID'].fillna(''),
    'status_date': df['status_date_str'],
    'sched_start': df['sched_start_str'],
    'contact_number': df['Contact Number'].fillna(''),
    'measurement': df['Measurement'].fillna(''),
    'measurement_date': df['measurement_date_str'],
    'measurement_result': df['Measurement Result'].fillna(''),
    'wo_class': df['WO Class'].fillna(''),
    'no_kontrak': df['No. Kontrak(KB/KL/P8)'].fillna(''),
    'product_type': df['Product Type'].fillna(''),
    'booking_date': df['booking_date_str'],
    'area_tif': df['Area TIF'].fillna(''),
    'district_tif': df['District TIF'].fillna(''),
    'regional_tif': df['Regional TIF'].fillna(''),
    'order_id_tsel': df['Order ID TSEL'].fillna(''),
    'channel_id_tsel': df['Channle ID TSEL'].fillna(''),
    'segment': df['Segment'],
    'is_ps': df['Is_PS'],
    'ps_duration_days': df['PS_Duration_Days'],
    'active_duration_days': df['Active_Duration_Days'],
    'is_ps_last_month': df['Is_PS_Last_Month']
})

db_df.to_sql('orders', conn, if_exists='replace', index=False)

cur.execute("CREATE INDEX idx_segment ON orders(segment);")
cur.execute("CREATE INDEX idx_crm_type ON orders(crm_order_type);")
cur.execute("CREATE INDEX idx_status ON orders(status);")
cur.execute("CREATE INDEX idx_is_ps ON orders(is_ps);")
cur.execute("CREATE INDEX idx_witel ON orders(witel);")
cur.execute("CREATE INDEX idx_sc_order ON orders(sc_order_no);")
conn.commit()

print("Step 4: Pre-building JSON summary stats...")
type_summary = []
for order_type, group in df.groupby('Tipe_Transaksi_Std'):
    total_orders = len(group)
    ps_count = int(group['Is_PS'].sum())
    ps_last_month = int(group['Is_PS_Last_Month'].sum())
    
    max_active_days = group['Active_Duration_Days'].max()
    max_ps_days = group['PS_Duration_Days'].max()
    avg_ps_days = group['PS_Duration_Days'].mean()
    
    type_summary.append({
        'tipe_transaksi': order_type,
        'total_order': total_orders,
        'total_ps': ps_count,
        'avg_ps_days': float(avg_ps_days) if pd.notna(avg_ps_days) else None,
        'avg_ps_hours': float(avg_ps_days * 24.0) if pd.notna(avg_ps_days) else None,
        'max_active_days': float(max_active_days) if pd.notna(max_active_days) else None,
        'max_ps_days': float(max_ps_days) if pd.notna(max_ps_days) else None,
        'ps_last_month': ps_last_month
    })

seg_summary = []
for seg, group in df.groupby('Segment'):
    total_orders = len(group)
    ps_count = int(group['Is_PS'].sum())
    ps_last_month = int(group['Is_PS_Last_Month'].sum())
    avg_ps_days = group['PS_Duration_Days'].mean()
    max_active_days = group['Active_Duration_Days'].max()
    
    seg_summary.append({
        'segment': seg,
        'total_order': total_orders,
        'total_ps': ps_count,
        'avg_ps_days': float(avg_ps_days) if pd.notna(avg_ps_days) else None,
        'avg_ps_hours': float(avg_ps_days * 24.0) if pd.notna(avg_ps_days) else None,
        'max_active_days': float(max_active_days) if pd.notna(max_active_days) else None,
        'ps_last_month': ps_last_month
    })

recent_df = df[df['Date Created'] >= one_month_ago].copy()
recent_df['date_key'] = recent_df['Date Created'].dt.strftime('%Y-%m-%d')
daily_counts = recent_df.groupby('date_key').agg(
    total_order=('Is_PS', 'count'),
    ps_count=('Is_PS', 'sum')
).reset_index().to_dict(orient='records')

summary_data = {
    'max_date': max_date.strftime('%Y-%m-%d %H:%M:%S'),
    'one_month_ago': one_month_ago.strftime('%Y-%m-%d %H:%M:%S'),
    'total_order_semesta': int(len(df)),
    'total_ps': int(df['Is_PS'].sum()),
    'ps_percentage': float((df['Is_PS'].sum() / len(df)) * 100.0),
    'total_ps_last_month': int(df['Is_PS_Last_Month'].sum()),
    'type_summary': type_summary,
    'segment_summary': seg_summary,
    'daily_trend': daily_counts
}

with open(json_summary_path, 'w', encoding='utf-8') as f:
    json.dump(summary_data, f, indent=2)

conn.close()
print("Done! All 31 columns stored into SQLite database.")
