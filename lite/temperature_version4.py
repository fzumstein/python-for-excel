import numpy as np
from xlwings import arg, func


@func
@arg("degrees", ndim=2)  # Apply the option ndim=2 to the argument "degrees"
def convert_to_celsius(degrees: np.ndarray, source: str = "fahrenheit"):
    if source.lower() == "fahrenheit":
        return (degrees - 32) * (5 / 9)
    elif source.lower() == "kelvin":
        return degrees - 273.15
    else:
        raise ValueError(f"Don't know how to convert from {source}")
