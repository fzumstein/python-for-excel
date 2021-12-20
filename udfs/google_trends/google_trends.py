import pandas as pd
from pytrends.request import TrendReq
import matplotlib.pyplot as plt
import xlwings as xw


@xw.func(call_in_wizard=False)
@xw.arg("mids", doc="Machine IDs: A range of max 5 cells")
@xw.arg("start_date", doc="A date-formatted cell")
@xw.arg("end_date", doc="A date-formatted cell")
def get_interest_over_time(mids, start_date, end_date):
    """Query Google Trends - replaces the Machine ID (mid) of
    common programming languages with their human-readable
    equivalent in the return value, e.g., instead of "/m/05z1_"
    it returns "Python".
    """
    # Check and transform parameters
    assert len(mids) <= 5, "Too many mids (max: 5)"
    start_date = start_date.date().isoformat()
    end_date = end_date.date().isoformat()

    # Make the Google Trends request and return the DataFrame
    trend = TrendReq(timeout=10)
    trend.build_payload(kw_list=mids,
                        timeframe=f"{start_date} {end_date}")
    df = trend.interest_over_time()

    # Replace Google's "mid" with a human-readable word
    mids = {"/m/05z1_": "Python", "/m/02p97": "JavaScript",
            "/m/0jgqg": "C++", "/m/07sbkfb": "Java", "/m/060kv": "PHP"}
    df = df.rename(columns=mids)

    # Drop the isPartial column
    return df.drop(columns="isPartial")


@xw.func
@xw.arg("df", pd.DataFrame)
def plot(df, name, caller):
    plt.style.use("seaborn")
    if not df.empty:
        caller.sheet.pictures.add(df.plot().get_figure(),
                                  top=caller.offset(row_offset=1).top,
                                  left=caller.left,
                                  name=name, update=True)
    return f"<Plot: {name}>"


if __name__ == "__main__":
    xw.serve()
