from xlwings import arg, func


@func
def convert_to_celsius(degrees, source="fahrenheit"):
    if source.lower() == "fahrenheit":
        return (degrees - 32) * (5 / 9)
    elif source.lower() == "kelvin":
        return degrees - 273.15
    else:
        raise ValueError(f"Don't know how to convert from {source}")
