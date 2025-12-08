from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule
from openpyxl.packaging.extended import ExtendedProperties
from openpyxl.styles import Alignment, Font

# Unfortunately, the latest version of OpenPyXL has an issue with
# chart formatting. This patch sets the Application property to fix it.
_original_init = ExtendedProperties.__init__


def patched_init(self, *args, **kwargs):
    _original_init(self, *args, **kwargs)
    self.Application = "Microsoft Excel"


ExtendedProperties.__init__ = patched_init

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

# DataFrame position and number of rows/columns
# openpxyl uses 1-based indices
startrow, startcol = 3, 2
nrows, ncols = summary.shape

with pd.ExcelWriter(
    this_dir / "sales_report_openpyxl.xlsx",
    engine="openpyxl",
) as writer:
    # pandas uses 0-based indices
    summary.to_excel(
        writer, sheet_name="Sheet1", startrow=startrow - 1, startcol=startcol - 1
    )

    # Get openpyxl book and sheet object
    book = writer.book
    sheet = writer.sheets["Sheet1"]

    # Set title
    sheet.cell(row=1, column=startcol, value="Sales Report")
    sheet.cell(row=1, column=startcol).font = Font(size=24, bold=True)

    # Sheet formatting
    sheet.sheet_view.showGridLines = False

    # Format the DataFrame with
    # - number format
    # - column width
    # - conditional formatting
    for row in range(startrow + 1, startrow + nrows + 1):
        for col in range(startcol + 1, startcol + ncols + 1):
            anchor_cell = sheet.cell(row=row, column=col)
            anchor_cell.number_format = "#,##0"
            anchor_cell.alignment = Alignment(horizontal="right")

    for anchor_cell in sheet["B"]:
        anchor_cell.number_format = "mmm yy"

    for col in range(startcol, startcol + ncols + 1):
        anchor_cell = sheet.cell(row=startrow, column=col)
        sheet.column_dimensions[anchor_cell.column_letter].width = 14

    first_cell = sheet.cell(row=startrow + 1, column=startcol + 1)
    last_cell = sheet.cell(row=startrow + nrows, column=startcol + ncols)
    range_address = f"{first_cell.coordinate}:{last_cell.coordinate}"
    sheet.conditional_formatting.add(
        range_address,
        CellIsRule(
            operator="lessThan",
            formula=["20000"],
            stopIfTrue=True,
            font=Font(color="E93423"),
        ),
    )

    # Chart
    chart = BarChart()
    chart.type = "col"
    chart.title = "Sales per Month and Store"
    chart.height = 11.5
    chart.width = 22.5
    chart.legend.overlay = False

    # Add each column as a series, ignoring total row and col
    data = Reference(
        sheet,
        min_col=startcol + 1,
        min_row=startrow,
        max_row=startrow + nrows - 1,
        max_col=startcol + ncols - 1,
    )
    categories = Reference(
        sheet, min_col=startcol, min_row=startrow + 1, max_row=startrow + nrows - 1
    )
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    anchor_cell = sheet.cell(row=startrow + nrows + 2, column=startcol)
    sheet.add_chart(chart=chart, anchor=anchor_cell.coordinate)

    # Chart formatting
    chart.y_axis.title = "Sales"
    chart.x_axis.title = summary.index.name
    # Hide y-axis line: spPR stands for ShapeProperties
    chart.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
