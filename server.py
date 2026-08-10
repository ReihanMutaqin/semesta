import http.server
import socketserver
import json
import sqlite3
import urllib.parse
import os
import sys
import subprocess

PORT = 3000
DB_PATH = r'C:\PROJEK\Data Semesta\semesta.db'
SUMMARY_PATH = r'C:\PROJEK\Data Semesta\summary.json'
EXCEL_PATH = r'C:\PROJEK\Data Semesta\export.xlsx'
CONVERT_SCRIPT = r'C:\PROJEK\Data Semesta\convert_to_sqlite.py'
STATIC_DIR = r'C:\PROJEK\Data Semesta\public'

class SemestaRequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=STATIC_DIR, **kwargs)

    def do_GET(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        query_params = urllib.parse.parse_qs(parsed_url.query)

        if path == '/api/summary':
            self.handle_api_summary()
        elif path == '/api/orders':
            self.handle_api_orders(query_params)
        elif path == '/api/witels':
            self.handle_api_witels()
        else:
            super().do_GET()

    def do_POST(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path

        if path == '/api/upload':
            self.handle_api_upload()
        else:
            self.send_error(404, "Endpoint not found")

    def handle_api_upload(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            if content_length == 0:
                self.send_json_response({'error': 'File payload is empty'}, status=400)
                return

            print(f"Receiving uploaded file ({content_length} bytes)...")

            # Chunked read for 25MB+ Excel files
            remaining = content_length
            file_data = bytearray()
            chunk_size = 65536
            while remaining > 0:
                chunk = self.rfile.read(min(remaining, chunk_size))
                if not chunk:
                    break
                file_data.extend(chunk)
                remaining -= len(chunk)

            # Save payload to export.xlsx
            with open(EXCEL_PATH, 'wb') as f:
                f.write(file_data)

            print(f"File saved to {EXCEL_PATH}. Running re-indexing script...")

            # Run convert_to_sqlite.py to regenerate semesta.db and summary.json
            proc = subprocess.run([sys.executable, CONVERT_SCRIPT], capture_output=True, text=True)

            if proc.returncode != 0:
                print("Error during re-indexing:", proc.stderr)
                self.send_json_response({'error': f'Failed to process excel file: {proc.stderr}'}, status=500)
                return

            print("Re-indexing completed successfully!")

            # Load updated summary
            if os.path.exists(SUMMARY_PATH):
                with open(SUMMARY_PATH, 'r', encoding='utf-8') as f:
                    updated_summary = json.load(f)
                self.send_json_response({
                    'success': True,
                    'message': 'File excel berhasil diupload dan diproses!',
                    'summary': updated_summary
                })
            else:
                self.send_json_response({'success': True, 'message': 'Data berhasil diproses!'})

        except Exception as e:
            print("Upload exception:", str(e))
            self.send_json_response({'error': str(e)}, status=500)

    def handle_api_summary(self):
        try:
            if os.path.exists(SUMMARY_PATH):
                with open(SUMMARY_PATH, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.send_json_response(data)
            else:
                self.send_error(404, "Summary file not found")
        except Exception as e:
            self.send_json_response({'error': str(e)}, status=500)

    def handle_api_orders(self, query_params):
        try:
            page = int(query_params.get('page', ['1'])[0])
            limit = int(query_params.get('limit', ['20'])[0])
            limit = min(max(limit, 1), 100)
            offset = (page - 1) * limit

            segment = query_params.get('segment', [''])[0]
            crm_type = query_params.get('crm_type', [''])[0]
            status = query_params.get('status', [''])[0]
            witel = query_params.get('witel', [''])[0]
            search = query_params.get('search', [''])[0].strip()
            sort_by = query_params.get('sort_by', ['date_created'])[0]
            sort_order = query_params.get('sort_order', ['DESC'])[0].upper()

            if sort_order not in ['ASC', 'DESC']:
                sort_order = 'DESC'

            allowed_sorts = {
                'date_created': 'date_created',
                'status_date': 'status_date',
                'sc_order_no': 'sc_order_no',
                'ps_duration_days': 'ps_duration_days',
                'active_duration_days': 'active_duration_days',
                'crm_order_type': 'crm_order_type',
                'status': 'status',
                'segment': 'segment'
            }
            sort_col = allowed_sorts.get(sort_by, 'date_created')

            where_clauses = []
            params = []

            if segment:
                where_clauses.append("segment = ?")
                params.append(segment)

            if crm_type:
                where_clauses.append("crm_order_type = ?")
                params.append(crm_type)

            if status:
                where_clauses.append("status = ?")
                params.append(status)

            if witel:
                where_clauses.append("witel = ?")
                params.append(witel)

            if search:
                search_like = f"%{search}%"
                where_clauses.append("(sc_order_no LIKE ? OR workorder LIKE ? OR service_no LIKE ? OR customer_name LIKE ? OR product_name LIKE ? OR description LIKE ?)")
                params.extend([search_like] * 6)

            where_sql = ""
            if where_clauses:
                where_sql = "WHERE " + " AND ".join(where_clauses)

            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()

            count_sql = f"SELECT COUNT(*) FROM orders {where_sql}"
            cur.execute(count_sql, params)
            total_records = cur.fetchone()[0]

            data_sql = f"""
                SELECT date_created, status_date, workorder, sc_order_no, service_no, 
                       crm_order_type, owner_group, status, product_name, customer_name, 
                       witel, segment, is_ps, ps_duration_days, active_duration_days, description
                FROM orders
                {where_sql}
                ORDER BY {sort_col} {sort_order} NULLS LAST
                LIMIT ? OFFSET ?
            """
            cur.execute(data_sql, params + [limit, offset])
            rows = [dict(row) for row in cur.fetchall()]

            total_pages = (total_records + limit - 1) // limit if total_records > 0 else 1

            conn.close()

            response_data = {
                'page': page,
                'limit': limit,
                'total_records': total_records,
                'total_pages': total_pages,
                'orders': rows
            }
            self.send_json_response(response_data)

        except Exception as e:
            self.send_json_response({'error': str(e)}, status=500)

    def handle_api_witels(self):
        try:
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            cur.execute("SELECT DISTINCT witel FROM orders WHERE witel != '' ORDER BY witel ASC")
            witels = [r[0] for r in cur.fetchall()]
            conn.close()
            self.send_json_response(witels)
        except Exception as e:
            self.send_json_response({'error': str(e)}, status=500)

    def send_json_response(self, data, status=200):
        body = json.dumps(data).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

if __name__ == '__main__':
    if not os.path.exists(STATIC_DIR):
        os.makedirs(STATIC_DIR)
    
    handler = SemestaRequestHandler
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), handler) as httpd:
        print(f"Server started at http://localhost:{PORT}")
        sys.stdout.flush()
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.")
