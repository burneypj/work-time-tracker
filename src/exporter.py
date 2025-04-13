from db import WorkSessionDB
import datetime
from utils import format_duration, increment_cell_row
from datetime import datetime
import xlwings as xw


def export_to_excel(excel_file, sheet_name, date_cell, start_cell, end_cell, duration_cell, date_based, db):
    """ Export session data to an Excel file using xlwings """
    sessions = db.get_sessions()

    if date_based:
        data = format_date_based_data(sessions)
    else:
        data = format_flat_data(sessions)

    # Open the workbook using xlwings
    app = xw.App(visible=False)  # Run Excel in the background
    try:
        wb = app.books.open(excel_file)
        ws = wb.sheets[sheet_name]

        # Write data to the specified cells
        for date, start_time, end_time, duration in data:
            if date_cell:
                ws.range(date_cell).value = date
                date_cell = increment_cell_row(date_cell)
            if start_cell:
                ws.range(start_cell).value = start_time
                start_cell = increment_cell_row(start_cell)
            if end_cell:
                ws.range(end_cell).value = end_time
                end_cell = increment_cell_row(end_cell)
            if duration_cell:
                ws.range(duration_cell).value = duration
                duration_cell = increment_cell_row(duration_cell)

        # Save the workbook
        wb.save()
    finally:
        wb.close()
        app.quit()


def format_flat_data(sessions):
    formatted = []
    for start_time_str, end_time_str, duration in sessions:
        start_time = datetime.fromisoformat(start_time_str)
        end_time = datetime.fromisoformat(end_time_str)

        date = start_time.date().isoformat()
        start_str = start_time.strftime("%H:%M:%S")
        end_str = end_time.strftime("%H:%M:%S")
        # Convert duration back to HH:MM:SS format
        hours, remainder = divmod(int(duration), 3600)
        minutes, seconds = divmod(remainder, 60)
        formatted_duration = f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"
        formatted.append((date, start_str, end_str, formatted_duration))
    return formatted


def format_date_based_data(sessions):
    """Format the data for date-based export."""
    from collections import defaultdict

    # Group sessions by date
    grouped_sessions = defaultdict(list)
    for session in sessions:
        start_time = datetime.fromisoformat(session[0])
        end_time = datetime.fromisoformat(session[1])
        duration_seconds = int(session[2])  # Duration is stored in seconds

        date = start_time.date().isoformat()
        grouped_sessions[date].append((start_time, end_time, duration_seconds))

    formatted_data = []

    for date, daily_sessions in grouped_sessions.items():
        # Skip rows with missing dates
        if not date:
            continue

        # Find the earliest start time, latest end time, and sum durations
        earliest_start = min(session[0] for session in daily_sessions)
        latest_end = max(session[1] for session in daily_sessions)
        total_duration_seconds = sum(session[2] for session in daily_sessions)

        # Convert total duration back to HH:MM:SS format
        hours, remainder = divmod(total_duration_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        formatted_duration = f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"

        formatted_data.append((date, earliest_start.strftime("%H:%M:%S"), latest_end.strftime("%H:%M:%S"), formatted_duration))

    return formatted_data
