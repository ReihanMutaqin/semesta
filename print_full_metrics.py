import sys
sys.path.append(r'C:\Users\reyha\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\site-packages')
import pandas as pd
import numpy as np

excel_path = r'C:\PROJEK\Data Semesta\export.xlsx'
df = pd.read_excel(excel_path)

df['Date Created'] = pd.to_datetime(df['Date Created'], errors='coerce')
df['Status Date'] = pd.to_datetime(df['Status Date'], errors='coerce')
max_date = max(df['Date Created'].max(), df['Status Date'].max())

# Standardize CRM Order Type
df['Tipe_Transaksi'] = df['CRM Order Type'].fillna('UNSPECIFIED').astype(str).str.upper().str.strip()
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
df['Is_PS'] = df['Status'].astype(str).str.upper().isin(ps_statuses)

# Duration calculation
df['PS_Duration_Days'] = np.where(
    df['Is_PS'] & df['Status Date'].notna() & df['Date Created'].notna(),
    (df['Status Date'] - df['Date Created']).dt.total_seconds() / 86400.0,
    np.nan
)
df['PS_Duration_Days'] = df['PS_Duration_Days'].apply(lambda x: x if x >= 0 else np.nan)

df['Active_Duration_Days'] = np.where(
    ~df['Is_PS'] & df['Date Created'].notna(),
    (max_date - df['Date Created']).dt.total_seconds() / 86400.0,
    np.nan
)

one_month_ago = max_date - pd.Timedelta(days=30)
df['Is_PS_Last_Month'] = df['Is_PS'] & (df['Status Date'] >= one_month_ago)

print("=== RINGKASAN PER TIPE TRANSAKSI (CRM ORDER TYPE STANDAR) ===")
summary_by_type = []
for order_type, group in df.groupby('Tipe_Transaksi_Std'):
    total_orders = len(group)
    ps_count = group['Is_PS'].sum()
    ps_last_month = group['Is_PS_Last_Month'].sum()
    
    max_active_days = group['Active_Duration_Days'].max()
    max_active_str = f"{max_active_days:.1f} hari" if pd.notna(max_active_days) else "-"
    
    max_ps_days = group['PS_Duration_Days'].max()
    max_ps_str = f"{max_ps_days:.1f} hari" if pd.notna(max_ps_days) else "-"
    
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
