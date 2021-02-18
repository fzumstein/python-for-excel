from pathlib import Path

import pandas as pd


# Directory of this file
this_dir = Path(__file__).resolve().parent

# Read in all Excel files from all subfolders of sales_data
parts = []
for path in (this_dir / "sales_data").rglob("*.xls*"):
    print(f'Reading {path.name}')
    part = pd.read_excel(path, index_col="transaction_id")
    parts.append(part)

# Combine the DataFrames from each file into a single DataFrame
# pandas takes care of properly aligning the columns
df = pd.concat(parts)

# Pivot each store into a column and sum up all transactions per date
pivot = pd.pivot_table(df,
                       index="transaction_date", columns="store",
                       values="amount", aggfunc="sum")

# Resample to end of month and assign an index name
summary = pivot.resample("M").sum()
summary.index.name = "Month"

# Write summary report to Excel file
summary.to_excel(this_dir / "sales_report_pandas.xlsx")
