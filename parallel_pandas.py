import multiprocessing
from itertools import repeat

import pandas as pd
import openpyxl


def _read_sheet(filename, sheet_name):
    # The leading underscore in the function name is used by convention
    # to mark it as "private", i.e., it shouldn't be used directly outside
    # of this module.
    df = pd.read_excel(filename, sheet_name=sheet_name, engine='openpyxl')
    return sheet_name, df


def read_excel(filename, sheet_name=None):
    if sheet_name is None:
        book = openpyxl.load_workbook(filename,
                                      read_only=True, data_only=True)
        sheet_name = book.sheetnames
        book.close()
    with multiprocessing.Pool() as pool:
        # By default, Pool spawns as many processes as there are CPU cores.
        # starmap maps a tuple of arguments to a function. The zip expression
        # produces a list with tuples of the following form:
        # [('filename.xlsx', 'Sheet1'), ('filename.xlsx', 'Sheet2)]
        data = pool.starmap(_read_sheet, zip(repeat(filename), sheet_name))
    return {i[0]: i[1] for i in data}
