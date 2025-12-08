import numpy as np
from xlwings import arg, func


@func
def revenue(base_fee, users, price):
    return base_fee + users * price


@func
@arg("users", np.array, ndim=2)
@arg("price", np.array)
def revenue2(base_fee, users, price):
    return base_fee + users * price
