import sys
sys.path.append(r'C:\Users\reyha\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\site-packages')
import pandas as pd
import numpy as np

excel_path = r'C:\PROJEK\Data Semesta\export.xlsx'
df = pd.read_excel(excel_path)

# Convert dates
df['Date Created'] = pd.to_datetime(df['Date Created'], errors='coerce')
df['Status Date'] = pd.to_datetime(df['Status Date'], errors='coerce')
df['Date Modified'] = pd.to_datetime(df['Date Modified'], errors='coerce')

# Reference date (max Date Created or max Status Date in dataset)
max_date = max(df['Date Created'].max(), df['Status Date'].max())

# Normalize CRM Order Type
df['CRM_Order_Type_Clean'] = df['CRM Order Type'].fillna('UNSPECIFIED').astype(str).str.upper().str.strip()
# Map 'NEW INSTALL' -> 'CREATE' if appropriate, or keep as standard
crm_map = {
    'NEW INSTALL': 'CREATE / NEW INSTALL',
    'CREATE': 'CREATE / NEW INSTALL',
    'MODIFY': 'MODIFY',
    'DISCONNECT': 'DISCONNECT',
    'SUSPEND': 'SUSPEND',
    'RESUME': 'RESUME',
    'MIGRATE': 'MIGRATE',
    'MOVE': 'MOVE',
    'UNSPECIFIED': 'UNSPECIFIED'
}
df['Tipe_Transaksi'] = df['CRM_Order_Type_Clean'].map(lambda x: crm_map.get(x, x))

# Determine PS Status (COMPLETED / COMPWORK / INSTCOMP / DEINSTCOMP / ACTCOMP / VALCOMP)
ps_statuses = ['COMPLETE', 'COMPWORK', 'INSTCOMP', 'DEINSTCOMP', 'ACTCOMP', 'VALCOMP']
df['Is_PS'] = df['Status'].astype(str).str.upper().isin(ps_statuses)

# Calculate PS Duration (in days)
df['PS_Duration_Days'] = np.where(
    df['Is_PS'] & df['Status Date'].notna() & df['Date Created'].notna(),
    (df['Status Date'] - df['Date Created']).dt.total_seconds() / 86400.0,
    np.nan
)
# Filter out invalid negative durations if any
df['PS_Duration_Days'] = df['PS_Duration_Days'].apply(lambda x: x if x >= 0 else np.nan)

# Active/Pending Duration (for incomplete orders)
df['Active_Duration_Days'] = np.where(
    ~df['Is_PS'] & df['Date Created'].notna(),
    (max_date - df['Date Created']).dt.total_seconds() / 86400.0,
    np.nan
)

# PS in the last 1 month (30 days) from max_date
one_month_ago = max_date - pd.Timedelta(days=30)
df['Is_PS_Last_Month'] = df['Is_PS'] & (df['Status Date'] >= one_month_ago)

print("==========================================")
print("       DATA SEMESTA METRICS SUMMARY       ")
print("==========================================")
print(f"Reference Max Date in Data: {max_date}")
print(f"1 Month Prior Threshold: {one_month_ago.strftime('%Y-%m-%d %H:%M:%S')}")
print(f"Total Order Semesta: {len(df):,}")
print(f"Total Completed (PS): {df['Is_PS'].sum():,} ({df['Is_PS'].mean()*100:.2f}%)")
print(f"Total PS 1 Bulan Kebelakang: {df['Is_PS_Last_Month'].sum():,}")

print("\n-----------------------------------------------------------------------------------------")
print(" METRICS PER TIPE TRANSAKSI (CRM ORDER TYPE) ")
print("-----------------------------------------------------------------------------------------")

summary_by_type = []
for order_type, group in df.groupby('Tipe_Transaksi'):
    total_orders = len(group)
    ps_count = group['Is_PS'].sum()
    ps_last_month = group['Is_PS_Last_Month'].sum()
    
    # Order Terlama (Active/Pending)
    max_active_days = group['Active_Duration_Days'].max()
    max_active_str = f"{max_active_days:.1f} hari" if pd.notna(max_active_days) else "-"
    
    # Order Terlama PS (Execution time max)
    max_ps_days = group['PS_Duration_Days'].max()
    max_ps_str = f"{max_ps_days:.1f} hari" if pd.notna(max_ps_days) else "-"
    
    # Average PS Duration
    avg_ps_days = group['PS_Duration_Days'].mean()
    avg_ps_str = f"{avg_ps_days:.2f} hari ({avg_ps_days*24:.1f} jam)" if pd.notna(avg_ps_days) else "-"
    
    summary_by_type.append({
        'Tipe Transaksi': order_type,
        'Total Order': total_orders,
        'Total PS': ps_count,
        'Rata-rata Durasi PS': avg_ps_str,
        'Order Pending Terlama': max_active_str,
        'Durasi PS Terlama': max_ps_str,
        'PS 1 Bulan Kebelakang': ps_last_month
    })

summary_df = pd.DataFrame(summary_by_type)
print(summary_df.to_string(index=False))

# Now let's classify by Segment: Modoroso, PDA HSI, IndiHome, Others
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

print("\n-----------------------------------------------------------------------------------------")
print(" METRICS PER SEGMENT (Modoroso, PDA HSI, IndiHome, Enterprise) ")
print("-----------------------------------------------------------------------------------------")

seg_summary = []
for seg, group in df.groupby('Segment'):
    total_orders = len(group)
    ps_count = group['Is_PS'].sum()
    ps_last_month = group['Is_PS_Last_Month'].sum()
    avg_ps_days = group['PS_Duration_Days'].mean()
    avg_ps_str = f"{avg_ps_days:.2f} hari ({avg_ps_days*24:.1f} jam)" if pd.notna(avg_ps_days) else "-"
    max_active_days = group['Active_Duration_Days'].max()
    max_active_str = f"{max_active_days:.1f} hari" if pd.notna(max_active_days) else "-"
    
    seg_summary.append({
        'Segment': seg,
        'Total Order': total_orders,
        'Total PS': ps_count,
        'Rata-rata Durasi PS': avg_ps_str,
        'Order Pending Terlama': max_active_str,
        'PS 1 Bulan Kebelakang': ps_last_month
    })

seg_df = pd.DataFrame(seg_summary)
print(seg_df.to_string(index=False))

