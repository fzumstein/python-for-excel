import numpy as np
import xlwings as xw


@xw.func
@xw.ret("raw")
def randn(i=1000, j=1000):
    """Returns an array with dimensions (i, j) with normally distributed
    pseudorandom numbers provided by NumPy's random.randn
    """
    return np.random.randn(i, j)
