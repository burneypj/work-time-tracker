from utils import format_duration, increment_cell_row
from datetime import datetime
import xlwings as xw
from PyQt5 import QtWidgets
from config import Config


class ExportConfigDialog(QtWidgets.QDialog):
    def __init__(self, excel_file, db, cfg, parent=None):
        super().__init__(parent)
        self.excel_file = excel_file
        self.db = db
        self.config = cfg
        self.workbook = None
        self.app = None

        self.setWindowTitle("Export Configuration")

        # Initialize UI elements for exporting
        self.sheet_name_combo = QtWidgets.QComboBox(self)
        self.load_workbook_and_sheets()

        self.date_cell_input = QtWidgets.QLineEdit(self)
        self.start_cell_input = QtWidgets.QLineEdit(self)
        self.end_cell_input = QtWidgets.QLineEdit(self)
        self.duration_cell_input = QtWidgets.QLineEdit(self)

        # Add a starting date input
        self.start_date_input = QtWidgets.QDateEdit(self)
        self.start_date_input.setCalendarPopup(True)
        self.start_date_input.setDate(datetime.now().date())
        # Set default values based on previous config or empty
        default_sheet = self.config.get('wb_sheet', '')
        if default_sheet:
            self.sheet_name_combo.setCurrentText(default_sheet)
        self.date_cell_input.setText(self.config.get('date_cell', 'A1'))
        self.start_cell_input.setText(self.config.get('start_cell', 'B1'))
        self.end_cell_input.setText(self.config.get('end_cell', 'C1'))
        self.duration_cell_input.setText(self.config.get('duration_cell', 'D1'))

        self.date_based_check = QtWidgets.QCheckBox("date-based export", self)
        self.date_based_check.setChecked(self.config.get('date_based', True))

        self.save_button = QtWidgets.QPushButton('Save', self)
        self.cancel_button = QtWidgets.QPushButton('Cancel', self)

        self.save_button.clicked.connect(self.save_export_settings)
        self.cancel_button.clicked.connect(self.reject)

        # Layout setup
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(QtWidgets.QLabel("Select Sheet:"))
        layout.addWidget(self.sheet_name_combo)
        layout.addWidget(QtWidgets.QLabel("Starting Date:"))
        layout.addWidget(self.start_date_input)
        layout.addWidget(QtWidgets.QLabel("Date Cell:"))
        layout.addWidget(self.date_cell_input)
        layout.addWidget(QtWidgets.QLabel("Start Cell:"))
        layout.addWidget(self.start_cell_input)
        layout.addWidget(QtWidgets.QLabel("End Cell:"))
        layout.addWidget(self.end_cell_input)
        layout.addWidget(QtWidgets.QLabel("Duration Cell:"))
        layout.addWidget(self.duration_cell_input)
        layout.addWidget(self.date_based_check)
        layout.addWidget(self.save_button)
        layout.addWidget(self.cancel_button)

        self.setLayout(layout)

    def load_workbook_and_sheets(self):
        """Open the workbook and load sheet names."""
        self.app = xw.App(visible=False)  # Run Excel in the background
        try:
            self.workbook = self.app.books.open(self.excel_file)
            sheet_names = [sheet.name for sheet in self.workbook.sheets]
            self.sheet_name_combo.addItems(sheet_names)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load workbook: {e}")
            self.reject()

    def save_export_settings(self):
        """Save the export settings and perform the export."""
        sheet_name = self.sheet_name_combo.currentText()
        date_cell = self.date_cell_input.text()
        start_cell = self.start_cell_input.text()
        end_cell = self.end_cell_input.text()
        duration_cell = self.duration_cell_input.text()
        date_based = self.date_based_check.isChecked()
        start_date = self.start_date_input.date().toPyDate()

        # Save these settings to config for future use
        self.config.set('wb_sheet', sheet_name)
        self.config.set('date_cell', date_cell)
        self.config.set('start_cell', start_cell)
        self.config.set('end_cell', end_cell)
        self.config.set('duration_cell', duration_cell)
        self.config.set('date_based', date_based)

        # Export to Excel
        self.write_to_excel(sheet_name, date_cell, start_cell, end_cell, duration_cell, date_based, start_date)
        self.close_excel()
        self.accept()  # Close the dialog after saving

    def write_to_excel(self, sheet_name, date_cell, start_cell, end_cell, duration_cell, date_based, start_date):
        """Export session data to an Excel file."""
        # Pass the start_date to get_sessions
        sessions = self.db.get_sessions(start_date=start_date if start_date else None)

        if date_based:
            data = self.format_date_based_data(sessions)
        else:
            data = self.format_flat_data(sessions)

        try:
            ws = self.workbook.sheets[sheet_name]

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
            self.workbook.save()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to write to workbook: {e}")

    def format_flat_data(self, sessions):
        formatted = []
        for start_time_str, end_time_str, duration in sessions:
            start_time = datetime.fromisoformat(start_time_str)
            end_time = datetime.fromisoformat(end_time_str)

            date = start_time.date().isoformat()
            start_str = start_time.strftime("%H:%M:%S")
            end_str = end_time.strftime("%H:%M:%S")
            formatted.append((date, start_str, end_str, format_duration(duration)))
        return formatted

    def format_date_based_data(self, sessions):
        """Format the data for date-based export, including placeholders for missing dates."""
        from collections import defaultdict
        from datetime import timedelta

        # Group sessions by date
        grouped_sessions = defaultdict(list)
        for session in sessions:
            start_time = datetime.fromisoformat(session[0])
            end_time = datetime.fromisoformat(session[1])
            duration_seconds = int(session[2])  # Duration is stored in seconds
            date = start_time.date().isoformat()
            grouped_sessions[date].append((start_time, end_time, duration_seconds))

        # Determine the full date range
        if sessions:
            first_date = datetime.fromisoformat(sessions[0][0]).date()
            last_date = datetime.fromisoformat(sessions[-1][0]).date()
        else:
            return []  # No sessions, return an empty list

        # Generate all dates in the range
        all_dates = [first_date + timedelta(days=i) for i in range((last_date - first_date).days + 1)]

        formatted_data = []
        for date in all_dates:
            date_str = date.isoformat()
            if date_str in grouped_sessions:
                # Process sessions for the date
                daily_sessions = grouped_sessions[date_str]
                earliest_start = min(session[0] for session in daily_sessions)
                latest_end = max(session[1] for session in daily_sessions)
                total_duration = format_duration(sum(session[2] for session in daily_sessions))
                formatted_data.append((date_str, earliest_start.strftime("%H:%M:%S"), latest_end.strftime("%H:%M:%S"), total_duration))
            else:
                # Add a placeholder for missing dates
                formatted_data.append((date_str, "", "", ""))

        return formatted_data

    def reject(self):
        self.close_excel()
        super().reject()

    def close_excel(self):
        """Close the Excel application."""
        if self.workbook:
            self.workbook.close()
            self.workbook = None
        if self.app:
            self.app.quit()
            self.app = None

def handle_excel_export(db, cfg):
    """Handle exporting session data to Excel."""
    excel_path = cfg.get('excel_path', '')
    options = QtWidgets.QFileDialog.Options()
    excel_file, _ = QtWidgets.QFileDialog.getOpenFileName(None, 'Select Excel File', excel_path, 'Excel Files (*.xlsx;*.xlsm)', options=options)

    if excel_file:
        cfg.set('excel_path', excel_file)
        # Open a dialog window for configuring the export settings
        dialog = ExportConfigDialog(excel_file, db, cfg)
        dialog.exec_()
