from pathlib import Path

import pandas as pd

# Directory of this file
this_dir = Path(__file__).resolve().parent

# Read in all files
parts = []
for path in (this_dir / "sales_data").rglob("*.xls*"):
    print(f"Reading {path.name}")
    part = pd.read_excel(path, engine="calamine")
    parts.append(part)

# Combine the DataFrames from each file into a single DataFrame
df = pd.concat(parts)

# Pivot each store into a column and sum up all transactions per date
pivot = pd.pivot_table(
    df, index="transaction_date", columns="store", values="amount", aggfunc="sum"
)

# Resample to end of month and assign an index name
summary = pivot.resample("ME").sum()
summary.index.name = "Month"

# Sort columns by total revenue
summary = summary.loc[:, summary.sum().sort_values().index]

# Add row and column totals
summary["Total"] = summary.sum(axis=1)
total_row = pd.DataFrame([summary.sum(axis=0)], index=["Total"])
summary = pd.concat([summary, total_row])

#### Write summary report to Excel file ####
summary.to_excel(this_dir / "sales_report_pandas.xlsx")
