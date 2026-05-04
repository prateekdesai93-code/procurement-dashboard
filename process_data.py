#!/usr/bin/env python3
"""
Purchase Order Data Processor
Reads an Excel PO file and outputs structured JSON for the dashboard.
Usage: python3 process_data.py --file YOUR_FILE.xlsx --output dashboard_data.json
"""

import pandas as pd
import json
import argparse
import sys
from pathlib import Path

def process_file(filepath: str) -> dict:
    df = pd.read_excel(filepath)

    # Normalise column names
    df.columns = [c.strip().lower() for c in df.columns]

    # Parse date
    date_col = next((c for c in df.columns if 'date' in c), None)
    if date_col:
        df['order date'] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['order date'])
    df['month_num'] = df['order date'].dt.month
    df['month_name'] = df['order date'].dt.strftime('%b')

    # Ensure required columns exist
    for col in ['product name', 'quantity', 'price', 'total', 'product code']:
        if col not in df.columns:
            df[col] = 0 if col in ['quantity', 'price', 'total'] else ''

    if 'supplier' not in df.columns:
        df['supplier'] = ''

    months_in_data = sorted(df['month_num'].unique())
    month_labels = {m: df[df['month_num']==m]['month_name'].iloc[0] for m in months_in_data}

    # ── KPIs ─────────────────────────────────────────────────────────
    kpis = {
        'total_spend':       round(float(df['total'].sum()), 2),
        'total_orders':      int(len(df)),
        'avg_order':         round(float(df['total'].mean()), 2),
        'unique_products':   int(df['product name'].nunique()),
        'unique_suppliers':  int(df['supplier'].dropna().nunique()),
        'date_range':        f"{df['order date'].min().strftime('%d %b %Y')} – {df['order date'].max().strftime('%d %b %Y')}",
        'months':            [month_labels[m] for m in months_in_data],
    }

    # ── Monthly spend ─────────────────────────────────────────────────
    monthly_spend = {
        'labels': [f"{month_labels[m]} {df[df['month_num']==m]['order date'].dt.year.mode()[0]}"
                   for m in months_in_data],
        'vals':   [round(float(df[df['month_num']==m]['total'].sum()), 2) for m in months_in_data],
    }

    # ── Daily spend ───────────────────────────────────────────────────
    daily = df.groupby(df['order date'].dt.date)['total'].sum().reset_index()
    daily = daily.sort_values('order date')
    daily_data = {
        'labels': [d.strftime('%d %b') for d in daily['order date']],
        'vals':   [round(float(v), 2) for v in daily['total']],
    }

    # ── Category spend ────────────────────────────────────────────────
    df['cat'] = df['product code'].str[:3]
    cat = df.groupby('cat')['total'].sum().sort_values(ascending=False).head(12)
    category_data = {
        'labels': list(cat.index),
        'vals':   [round(float(v), 2) for v in cat.values],
    }

    # ── Supplier spend ────────────────────────────────────────────────
    sup = df.dropna(subset=['supplier'])
    sup = sup[sup['supplier'].astype(str).str.strip() != '']
    sup_grp = sup.groupby('supplier')['total'].sum().sort_values(ascending=False)
    supplier_data = {
        'labels': list(sup_grp.index),
        'vals':   [round(float(v), 2) for v in sup_grp.values],
    }

    # ── All products with monthly breakdown ───────────────────────────
    all_products = df.groupby('product name').agg(
        total_qty=('quantity', 'sum'),
        total_spend=('total', 'sum'),
        order_count=('total', 'count'),
    ).reset_index().sort_values('total_qty', ascending=False)

    products = []
    for _, row in all_products.iterrows():
        prod = row['product name']
        sub = df[df['product name'] == prod]
        supplier_val = ''
        if 'supplier' in sub.columns:
            s = sub['supplier'].dropna()
            if len(s) > 0:
                supplier_val = str(s.iloc[0]) if pd.notna(s.iloc[0]) else ''

        entry = {
            'product':        prod,
            'code':           str(sub['product code'].iloc[0]) if len(sub) > 0 else '',
            'supplier':       supplier_val,
            'avg_price':      round(float(sub['price'].mean()), 4),
            'total_qty':      int(row['total_qty']),
            'total_spend':    round(float(row['total_spend']), 2),
            'order_count':    int(row['order_count']),
        }

        # Monthly qty and spend for each month in data
        for m in months_in_data:
            key = month_labels[m].lower()
            entry[f'{key}_qty']   = int(sub[sub['month_num'] == m]['quantity'].sum())
            entry[f'{key}_spend'] = round(float(sub[sub['month_num'] == m]['total'].sum()), 2)

        # Planning
        n = len(months_in_data)
        entry['avg_monthly_qty'] = round(entry['total_qty'] / n, 1)
        entry['suggested_next']  = max(round(entry['total_qty'] / n * 1.1), 0)

        # Trend: compare last month vs first month
        first_qty = entry.get(f"{month_labels[months_in_data[0]].lower()}_qty", 0)
        last_qty  = entry.get(f"{month_labels[months_in_data[-1]].lower()}_qty", 0)
        if last_qty > first_qty:
            entry['trend'] = 'up'
        elif last_qty < first_qty and last_qty > 0:
            entry['trend'] = 'down'
        else:
            entry['trend'] = 'stable'

        products.append(entry)

    return {
        'kpis':     kpis,
        'monthly':  monthly_spend,
        'daily':    daily_data,
        'category': category_data,
        'supplier': supplier_data,
        'products': products,
        'months':   [month_labels[m] for m in months_in_data],
    }


def main():
    parser = argparse.ArgumentParser(description='Process PO Excel file to JSON')
    parser.add_argument('--file',   required=True, help='Path to .xlsx file')
    parser.add_argument('--output', default='dashboard_data.json', help='Output JSON path')
    args = parser.parse_args()

    if not Path(args.file).exists():
        print(f"ERROR: File not found: {args.file}")
        sys.exit(1)

    print(f"Processing: {args.file}")
    data = process_file(args.file)
    print(f"  Products:  {len(data['products'])}")
    print(f"  Months:    {', '.join(data['months'])}")
    print(f"  Spend:     R{data['kpis']['total_spend']:,.2f}")

    with open(args.output, 'w') as f:
        json.dump(data, f, indent=2)
    print(f"  Saved: {args.output}")


if __name__ == '__main__':
    main()
