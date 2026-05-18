#!/usr/bin/env python3
"""Quick test of Yandex Wordstat API v2."""
import os
import json
import requests
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("YANDEX_API_KEY", "")
FOLDER_ID = os.getenv("FOLDER_ID", "")

if not API_KEY or not FOLDER_ID:
    print("ERROR: Missing API key or folder ID in .env")
    exit(1)

url = "https://searchapi.api.cloud.yandex.net/v2/wordstat/dynamics"
headers = {
    "Authorization": f"Api-Key {API_KEY}",
    "Content-Type": "application/json",
}

payload = {
    "folderId": FOLDER_ID,
    "phrase": "купить оборудование для",
    "period": "MONTH",  # FIXED: was MONTHLY
    "from_date": "2024-01-01T00:00:00Z",
    "to_date": "2024-03-01T00:00:00Z",
    "geo_ids": [225],
    "group_by": "TIME",
}

print(f"URL: {url}")
print(f"Payload: {json.dumps(payload, indent=2, ensure_ascii=False)}")
print()

resp = requests.post(url, headers=headers, json=payload, timeout=30)

print(f"Status: {resp.status_code}")
if resp.status_code == 200:
    data = resp.json()
    print(f"Response keys: {list(data.keys())}")
    print(f"Response (first 2000 chars):")
    print(json.dumps(data, indent=2, ensure_ascii=False)[:2000])
else:
    print(f"Response body: {resp.text[:1000]}")