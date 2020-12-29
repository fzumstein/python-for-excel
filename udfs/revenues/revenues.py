import numpy as np
import xlwings as xw


@xw.func
def revenue(base_fee, users, price):
    return base_fee + users * price


@xw.func
@xw.arg("users", np.array, ndim=2)
@xw.arg("price", np.array)
def revenue2(base_fee, users, price):
    return base_fee + users * price
