import re


def increment_cell_row(cell_ref):
    """
    Increments the row number in an Excel cell reference.
    E.g., 'A1' -> 'A2', 'B10' -> 'B11'
    """
    match = re.match(r"([A-Z]+)([0-9]+)", cell_ref, re.I)
    if match:
        col, row = match.groups()
        return f"{col}{int(row) + 1}"
    else:
        raise ValueError(f"Invalid cell reference: {cell_ref}")


def format_duration(seconds):
    hours, remainder = divmod(int(seconds), 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"
