# Purchase Order Dashboard

An interactive analytics dashboard for purchase order data, built with Python and Chart.js.

## Features
- Overview: KPIs, daily/monthly spend trends, top products by spend and quantity
- Products tab: searchable/sortable table with monthly breakdown per product
- Usage Charts: grouped and stacked monthly usage by product group
- Order Planning: editable next-month order quantities with CSV export

## Setup

```bash
pip install pandas openpyxl xlsxwriter
```

## Usage

**Step 1 — Process your Excel file:**
```bash
python3 process_data.py --file YOUR_FILE.xlsx --output dashboard_data.json
```

**Step 2 — Build the HTML dashboard:**
```bash
python3 build_dashboard.py --data dashboard_data.json --output index.html
```

Open `index.html` in any browser or host it on GitHub Pages.

**Optional — Build Excel comparison (2025 vs 2026):**
```bash
python3 build_excel.py --file2025 Book1.xlsx --file2026 FOR.xlsx --output Comparison.xlsx
```

## Excel File Format

Required columns: `order number`, `product code`, `order date` (DD/MM/YYYY), `product name`, `quantity`, `price`, `total`

Optional: `supplier`, `supplier product name`
