from xlwings import arg, func


@func
def convert_to_celsius(degrees, source="fahrenheit"):
    if not isinstance(degrees, list):
        # Make sure the function also works with single cells.
        # We'll see how to handle this with with the "ndim" option below.
        degrees = [degrees]

    results = []
    for degree in degrees:
        if source.lower() == "fahrenheit":
            results.append((degree - 32) * (5 / 9))
        elif source.lower() == "kelvin":
            results.append(degree - 273.15)
        else:
            raise ValueError(f"Don't know how to convert from {source}")

    return results
