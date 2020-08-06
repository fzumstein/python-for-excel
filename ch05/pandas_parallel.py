import multiprocessing
from itertools import repeat

import pandas as pd
import openpyxl


def _read_sheet(filename, sheet_name):
    df = pd.read_excel(filename, sheet_name=sheet_name, engine='openpyxl')
    return sheet_name, df

def read_excel(filename, sheet_name=None):
    if sheet_name is None:
        book = openpyxl.load_workbook(filename,
                                      read_only=True, data_only=True)
        sheet_name = book.sheetnames
        book.close()
    with multiprocessing.Pool() as pool:
        data = pool.starmap(_read_sheet, zip(repeat(filename), sheet_name))
    return {i[0]: i[1] for i in data}
