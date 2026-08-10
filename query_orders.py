import sys
import sqlite3
import json
import urllib.parse
import os

def query_orders():
    try:
        query_str = sys.argv[1] if len(sys.argv) > 1 else ""
        params = urllib.parse.parse_qs(query_str)

        page = int(params.get('page', ['1'])[0])
        limit = int(params.get('limit', ['20'])[0])
        limit = min(max(limit, 1), 100)
        offset = (page - 1) * limit

        segment = params.get('segment', [''])[0]
        crm_type = params.get('crm_type', [''])[0]
        status = params.get('status', [''])[0]
        witel = params.get('witel', [''])[0]
        search = params.get('search', [''])[0].strip()

        where_clauses = []
        sql_params = []

        if segment:
            where_clauses.append("segment = ?")
            sql_params.append(segment)

        if crm_type:
            where_clauses.append("crm_order_type = ?")
            sql_params.append(crm_type)

        if status:
            where_clauses.append("status = ?")
            sql_params.append(status)

        if witel:
            where_clauses.append("witel = ?")
            sql_params.append(witel)

        if search:
            search_like = f"%{search}%"
            where_clauses.append("(sc_order_no LIKE ? OR workorder LIKE ? OR service_no LIKE ? OR customer_name LIKE ? OR product_name LIKE ? OR description LIKE ? OR address LIKE ?)")
            sql_params.extend([search_like] * 7)

        where_sql = ""
        if where_clauses:
            where_sql = "WHERE " + " AND ".join(where_clauses)

        db_path = r'C:\PROJEK\Data Semesta\semesta.db'
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        count_sql = f"SELECT COUNT(*) FROM orders {where_sql}"
        cur.execute(count_sql, sql_params)
        total_records = cur.fetchone()[0]

        # Select ALL columns
        data_sql = f"""
            SELECT *
            FROM orders
            {where_sql}
            ORDER BY date_created DESC NULLS LAST
            LIMIT ? OFFSET ?
        """
        cur.execute(data_sql, sql_params + [limit, offset])
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()

        total_pages = (total_records + limit - 1) // limit if total_records > 0 else 1

        result = {
            'page': page,
            'limit': limit,
            'total_records': total_records,
            'total_pages': total_pages,
            'orders': rows
        }
        print(json.dumps(result))

    except Exception as e:
        print(json.dumps({'error': str(e)}))

if __name__ == '__main__':
    query_orders()
