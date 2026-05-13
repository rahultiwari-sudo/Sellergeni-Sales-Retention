#!/usr/bin/env python3
"""
Fetches the Google Sheets export (.xlsx), computes monthly aggregates grouped by
closure month and onboarding month, and writes CSV outputs to /tmp and repo folder.

Usage: python3 scripts/compute_monthly_aggregates.py
"""
from pathlib import Path
import re
import sys
import requests

SHEET_URL = "https://docs.google.com/spreadsheets/d/1ad_HCXJ7rQ5onzBLnhlfMKvBzD1QdSVhugttVjmu8GE/export?format=xlsx"
OUT_DIR = Path("./reports")
OUT_DIR.mkdir(parents=True, exist_ok=True)
TMP_PATH = Path("/tmp/sheet_live.xlsx")

try:
    import pandas as pd
except Exception:
    print("Installing pandas and openpyxl...", file=sys.stderr)
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas', 'openpyxl'])
    import pandas as pd


def to_number(x):
    try:
        if pd.isna(x):
            return 0.0
        s = str(x)
        s = re.sub(r"[^0-9.-]", "", s)
        return float(s) if s else 0.0
    except Exception:
        return 0.0


def parse_date(x):
    if x is None:
        return None
    s = str(x).strip()
    if not s:
        return None
    # numeric -> excel serial
    if re.match(r'^\d+(?:\.\d+)?$', s):
        try:
            val = float(s)
            from datetime import timedelta, datetime
            excel_epoch = datetime(1899, 12, 30)
            return excel_epoch + timedelta(days=val)
        except Exception:
            pass
    try:
        dt = pd.to_datetime(s, dayfirst=False, errors='coerce')
        if pd.isna(dt):
            return None
        return dt.to_pydatetime()
    except Exception:
        return None


def find_field_val(row, df_cols, keywords):
    lowmap = {c: re.sub(r"\s+", "", str(c)).lower() for c in df_cols}
    for key in keywords:
        norm = key.lower().replace(' ', '')
        for c, cl in lowmap.items():
            if norm in cl:
                return row.get(c, '')
    return ''


def compute_aggregates(xlsx_path):
    xl = pd.read_excel(xlsx_path, sheet_name=None, dtype=str)
    closure = {}
    onboard = {}

    for sheet_name, df in xl.items():
        if df is None or df.shape[0] == 0:
            continue
        cols = list(df.columns)
        for idx, r in df.iterrows():
            amount_pitched = find_field_val(r, cols, ['Amount Pitched', 'AmountPitched', 'amountpitched'])
            closure_amount = find_field_val(r, cols, ['Closure Amount', 'ClosureAmount', 'closureamount'])
            closure_date = find_field_val(r, cols, ['Closure Date', 'ClosureDate', 'closuredate'])
            onboarding_date = find_field_val(r, cols, ['Date of Onboarding', 'Onboarding Date', 'onboardingdate', 'Date'])

            ap = to_number(amount_pitched)
            ca = to_number(closure_amount)
            cd = parse_date(closure_date)
            od = parse_date(onboarding_date)

            if cd is not None:
                key = f"{cd.year:04d}-{cd.month:02d}"
                if key not in closure:
                    closure[key] = {'rows':0, 'pitched':0.0, 'closure':0.0}
                closure[key]['rows'] += 1
                closure[key]['pitched'] += ap
                closure[key]['closure'] += ca

            if od is not None:
                key2 = f"{od.year:04d}-{od.month:02d}"
                if key2 not in onboard:
                    onboard[key2] = {'rows':0, 'pitched':0.0, 'closure':0.0}
                onboard[key2]['rows'] += 1
                onboard[key2]['pitched'] += ap
                onboard[key2]['closure'] += ca

    return closure, onboard


def write_csv(dct, path):
    import csv
    rows = []
    for k in sorted(dct.keys()):
        rows.append((k, dct[k]['rows'], int(dct[k]['pitched']), int(dct[k]['closure'])))
    with open(path, 'w', newline='') as f:
        w = csv.writer(f)
        w.writerow(['month','rows','pitched','closure'])
        w.writerows(rows)
    return path


def main():
    print('Downloading sheet...')
    r = requests.get(SHEET_URL, timeout=60)
    r.raise_for_status()
    TMP_PATH.write_bytes(r.content)
    print('Computing aggregates...')
    closure, onboard = compute_aggregates(str(TMP_PATH))
    cpath = write_csv(closure, OUT_DIR / 'closure_by_month.csv')
    opath = write_csv(onboard, OUT_DIR / 'onboarding_by_month.csv')
    print('Wrote:', cpath, opath)
    print('\nClosure-month summary:')
    for k in sorted(closure.keys()):
        v = closure[k]
        print(f"{k}: rows={v['rows']}, pitched={int(v['pitched'])}, closure={int(v['closure'])}")
    print('\nOnboarding-month summary:')
    for k in sorted(onboard.keys()):
        v = onboard[k]
        print(f"{k}: rows={v['rows']}, pitched={int(v['pitched'])}, closure={int(v['closure'])}")

if __name__ == '__main__':
    main()
