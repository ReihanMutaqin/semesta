import sqlite3
import json
import os

db_path = r'C:\PROJEK\Data Semesta\semesta.db'
summary_path = r'C:\PROJEK\Data Semesta\summary.json'

print("Resetting database to initial empty state...")

# 1. Create Empty SQLite database
if os.path.exists(db_path):
    os.remove(db_path)

conn = sqlite3.connect(db_path)
cur = conn.cursor()

cur.execute("""
    CREATE TABLE orders (
        date_created TEXT,
        date_modified TEXT,
        workorder TEXT,
        sc_order_no TEXT,
        oss_order_id TEXT,
        service_no TEXT,
        description TEXT,
        crm_order_type TEXT,
        owner_group TEXT,
        status TEXT,
        product_name TEXT,
        address TEXT,
        witel TEXT,
        customer_name TEXT,
        workzone TEXT,
        region_site_id TEXT,
        status_date TEXT,
        sched_start TEXT,
        contact_number TEXT,
        measurement TEXT,
        measurement_date TEXT,
        measurement_result TEXT,
        wo_class TEXT,
        no_kontrak TEXT,
        product_type TEXT,
        booking_date TEXT,
        area_tif TEXT,
        district_tif TEXT,
        regional_tif TEXT,
        order_id_tsel TEXT,
        channel_id_tsel TEXT,
        segment TEXT,
        is_ps INTEGER,
        ps_duration_days REAL,
        active_duration_days REAL,
        is_ps_last_month INTEGER
    );
""")

cur.execute("CREATE INDEX idx_segment ON orders(segment);")
cur.execute("CREATE INDEX idx_crm_type ON orders(crm_order_type);")
cur.execute("CREATE INDEX idx_status ON orders(status);")
cur.execute("CREATE INDEX idx_is_ps ON orders(is_ps);")
cur.execute("CREATE INDEX idx_witel ON orders(witel);")
cur.execute("CREATE INDEX idx_sc_order ON orders(sc_order_no);")
conn.commit()
conn.close()

# 2. Create Empty summary.json
empty_summary = {
    'max_date': 'Belum Ada Data',
    'one_month_ago': '-',
    'total_order_semesta': 0,
    'total_ps': 0,
    'ps_percentage': 0.0,
    'total_ps_last_month': 0,
    'type_summary': [
        {'tipe_transaksi': 'CREATE', 'total_order': 0, 'total_ps': 0, 'avg_ps_days': None, 'avg_ps_hours': None, 'max_active_days': None, 'max_ps_days': None, 'ps_last_month': 0},
        {'tipe_transaksi': 'MODIFY', 'total_order': 0, 'total_ps': 0, 'avg_ps_days': None, 'avg_ps_hours': None, 'max_active_days': None, 'max_ps_days': None, 'ps_last_month': 0},
        {'tipe_transaksi': 'DISCONNECT', 'total_order': 0, 'total_ps': 0, 'avg_ps_days': None, 'avg_ps_hours': None, 'max_active_days': None, 'max_ps_days': None, 'ps_last_month': 0},
        {'tipe_transaksi': 'SUSPEND', 'total_order': 0, 'total_ps': 0, 'avg_ps_days': None, 'avg_ps_hours': None, 'max_active_days': None, 'max_ps_days': None, 'ps_last_month': 0},
        {'tipe_transaksi': 'MIGRATE', 'total_order': 0, 'total_ps': 0, 'avg_ps_days': None, 'avg_ps_hours': None, 'max_active_days': None, 'max_ps_days': None, 'ps_last_month': 0}
    ],
    'segment_summary': [
        {'segment': 'Modoroso', 'total_order': 0, 'total_ps': 0, 'avg_ps_days': None, 'avg_ps_hours': None, 'max_active_days': None, 'ps_last_month': 0},
        {'segment': 'PDA HSI', 'total_order': 0, 'total_ps': 0, 'avg_ps_days': None, 'avg_ps_hours': None, 'max_active_days': None, 'ps_last_month': 0},
        {'segment': 'IndiHome', 'total_order': 0, 'total_ps': 0, 'avg_ps_days': None, 'avg_ps_hours': None, 'max_active_days': None, 'ps_last_month': 0},
        {'segment': 'Enterprise / Lainnya', 'total_order': 0, 'total_ps': 0, 'avg_ps_days': None, 'avg_ps_hours': None, 'max_active_days': None, 'ps_last_month': 0}
    ],
    'daily_trend': []
}

with open(summary_path, 'w', encoding='utf-8') as f:
    json.dump(empty_summary, f, indent=2)

print("Reset complete! Initial state is now empty (0 orders).")
