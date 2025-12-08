import pandas as pd
import xlwings as xw
from xlwings import arg, func


@func(call_in_wizard=False)
@arg("tickers", ndim=1, doc="A range of tickers.")
@arg("start_date", doc="A date-formatted cell")
@arg("end_date", doc="A date-formatted cell")
@arg("rebase", doc="If TRUE, starts all time series at 100")
@arg("column", doc="Defines the column name. Default: adj_close")
def get_stock_history(
    tickers, start_date=None, end_date=None, rebase=False, column="adj_close"
):
    """Fetch historical stock data for given tickers and combine them into a
    single DataFrame.
    """
    base_url = "https://raw.githubusercontent.com/fzumstein/python-for-excel/2e/csv"
    parts = []

    # Fetch data for each ticker
    for ticker in tickers:
        # Download the data from the online GitHub repository
        url = f"{base_url}/{ticker}.csv"
        df = pd.read_csv(url, parse_dates=["date"], index_col="date")
        df.index.name = "Date"
        parts.append(df[[column]].rename(columns={column: ticker}))

    # Combine all DataFrames
    result = pd.concat(parts, axis=1)

    # Filter by date range
    if start_date is not None or end_date is not None:
        result = result.loc[start_date:end_date, :]

    # Rebase
    if rebase:
        result = result / result.iloc[0] * 100

    return result


@func
@arg("df", pd.DataFrame)
def plot(df, name, caller):
    if not df.empty:
        caller.sheet.pictures.add(
            df.plot().get_figure(),
            anchor=caller.offset(row_offset=1),
            name=name,
            update=True,
        )
    return f"<Plot: {name}>"


if __name__ == "__main__":
    xw.serve()
