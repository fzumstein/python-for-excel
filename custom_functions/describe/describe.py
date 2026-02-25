import pandas as pd
from xlwings import arg, func

@func
@arg("df", pd.DataFrame, index=True, header=True)
def describe(df):
    return df.describe()
