from typing import Annotated

import numpy as np
from xlwings import arg, func

Array2d = Annotated[np.ndarray, {"ndim": 2}]


@func
def convert_to_celsius(degrees: Array2d, source: str = "fahrenheit"):
    if source.lower() == "fahrenheit":
        return (degrees - 32) * (5 / 9)
    elif source.lower() == "kelvin":
        return degrees - 273.15
    else:
        raise ValueError(f"Don't know how to convert from {source}")
