import multiprocessing
from itertools import repeat

import xlrd
import excel


def _read_sheet(filename, sheetname):
    with xlrd.open_workbook(filename, on_demand=True) as book:
        sheet = book.sheet_by_name(sheetname)
        data = excel.read(sheet)
    return sheet.name, data

def open_workbook(filename, sheetnames=None):
    if sheetnames is None:
        with xlrd.open_workbook(filename, on_demand=True) as book:
            sheetnames = book.sheet_names()
    with multiprocessing.Pool() as pool:
        data = pool.starmap(_read_sheet, zip(repeat(filename), sheetnames))
    return {i[0]: i[1] for i in data}
