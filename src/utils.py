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
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    seconds = int(seconds % 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"
