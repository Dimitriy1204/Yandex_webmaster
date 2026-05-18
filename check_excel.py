#!/usr/bin/env python3
"""Check yandex_analysis.xlsx content."""
import os
import pandas as pd

xl = pd.ExcelFile('yandex_analysis.xlsx')
print('Sheets:', xl.sheet_names)
for s in xl.sheet_names:
    df = pd.read_excel(xl, s)
    print(f'\n--- {s} ---')
    print(f'  Shape: {df.shape}')
    print(f'  Columns: {list(df.columns)}')
    if df.shape[0] > 0:
        print(f'  First row: {dict(df.iloc[0])}')
print(f'\nFile size: {os.path.getsize("yandex_analysis.xlsx") / 1024:.1f} KB')