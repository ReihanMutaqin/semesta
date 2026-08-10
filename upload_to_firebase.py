"""
Script Pengunggah Data Semesta ke Firebase Realtime Database / Firestore.
Jalankan script ini secara lokal untuk mengimpor file summary.json & semesta.db ke Firebase Cloud.
"""

import urllib.request
import json
import os
import sys

# Isikan URL Firebase Realtime Database milik Anda di sini
FIREBASE_DATABASE_URL = os.environ.get("FIREBASE_DATABASE_URL", "https://YOUR_PROJECT_ID-default-rtdb.asia-southeast1.firebasedatabase.app")
SUMMARY_JSON_PATH = r"C:\PROJEK\Data Semesta\summary.json"

def upload_summary_to_firebase():
    if not os.path.exists(SUMMARY_JSON_PATH):
        print(f"Error: {SUMMARY_JSON_PATH} tidak ditemukan. Jalankan convert_to_sqlite.py terlebih dahulu.")
        return

    print("Step 1: Reading summary.json...")
    with open(SUMMARY_JSON_PATH, 'r', encoding='utf-8') as f:
        summary_data = json.load(f)

    print("Step 2: Uploading summary metrics to Firebase Realtime Database...")
    url = f"{FIREBASE_DATABASE_URL}/summary.json"
    req_body = json.dumps(summary_data).encode('utf-8')

    req = urllib.request.Request(
        url,
        data=req_body,
        headers={'Content-Type': 'application/json'},
        method='PUT'
    )

    try:
        with urllib.request.urlopen(req) as resp:
            print("Summary successfully uploaded to Firebase RTDB! Status:", resp.status)
    except Exception as e:
        print("Error uploading summary to Firebase:", e)

if __name__ == '__main__':
    upload_summary_to_firebase()
