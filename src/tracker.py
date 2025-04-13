from PyQt5 import QtWidgets, QtCore
import datetime
from db import WorkSessionDB
from exporter import export_to_excel
from config import Config


class TimeTrackerApp(QtWidgets.QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db  # WorkSessionDB instance for database access
        self.start_time = None
        self.end_time = None
        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.update_time)

        # UI elements
        self.init_ui()

    def init_ui(self):
        # Set up the UI for session tracking
        self.setWindowTitle('Time Tracker')
        self.setGeometry(100, 100, 300, 200)

        self.start_button = QtWidgets.QPushButton('Start', self)
        self.stop_button = QtWidgets.QPushButton('Stop', self)
        self.export_button = QtWidgets.QPushButton('Export to Excel', self)

        # Add a label to display the running duration
        self.duration_label = QtWidgets.QLabel("Duration: 00:00:00", self)
        self.duration_label.setAlignment(QtCore.Qt.AlignCenter)

        self.start_button.clicked.connect(self.start_session)
        self.stop_button.clicked.connect(self.stop_session)
        self.export_button.clicked.connect(self.export_to_excel)

        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)

        # Layout setup
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self.duration_label)  # Add the duration label to the layout
        layout.addWidget(self.start_button)
        layout.addWidget(self.stop_button)
        layout.addWidget(self.export_button)
        self.setLayout(layout)

    def start_session(self):
        """ Start the session """
        self.start_time = datetime.datetime.now()
        self.timer.start(1000)  # Update every second
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)

    def stop_session(self):
        """ Stop the session and save data """
        self.end_time = datetime.datetime.now()
        total_seconds = int((self.end_time - self.start_time).total_seconds())  # Round down to the nearest second
        # Log the session to the database (store duration in seconds)
        self.db.add_session(self.start_time, self.end_time, total_seconds)
        # Stop the timer and reset buttons
        self.timer.stop()
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)

    def update_time(self):
        """ Update the UI with the current time """
        if self.start_time:
            elapsed_time = datetime.datetime.now() - self.start_time
            hours, remainder = divmod(elapsed_time.total_seconds(), 3600)
            minutes, seconds = divmod(remainder, 60)
            self.duration_label.setText(f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}")

    def export_to_excel(self):
        """ Handle exporting session data to Excel """
        # Open file dialog to choose the Excel file
        config = Config()
        excel_path = config.get('excel_path', '')
        options = QtWidgets.QFileDialog.Options()
        excel_file, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Select Excel File', excel_path, 'Excel Files (*.xlsx;*.xlsm)', options=options)

        if excel_file:
            config.set('excel_path', excel_file)
            """ Open a dialog window for configuring the export settings """
            dialog = ExportConfigDialog(excel_file, self.db, config)
            dialog.exec_()


class ExportConfigDialog(QtWidgets.QDialog):
    def __init__(self, excel_file, db, cfg, parent=None):
        super().__init__(parent)
        self.excel_file = excel_file
        self.db = db
        self.config = cfg

        self.setWindowTitle("Export Configuration")

        # Initialize UI elements for exporting
        self.sheet_name_combo = QtWidgets.QComboBox(self)
        self.sheet_name_combo.addItems(self.get_sheet_names())

        self.date_cell_input = QtWidgets.QLineEdit(self)
        self.start_cell_input = QtWidgets.QLineEdit(self)
        self.end_cell_input = QtWidgets.QLineEdit(self)
        self.duration_cell_input = QtWidgets.QLineEdit(self)

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

    def get_sheet_names(self):
        """ Get sheet names from the Excel file """
        import openpyxl
        workbook = openpyxl.load_workbook(self.excel_file)
        return workbook.sheetnames

    def save_export_settings(self):
        """ Save the export settings and perform the export """
        sheet_name = self.sheet_name_combo.currentText()
        date_cell = self.date_cell_input.text()
        start_cell = self.start_cell_input.text()
        end_cell = self.end_cell_input.text()
        duration_cell = self.duration_cell_input.text()
        date_based = self.date_based_check.isChecked()

        # Save these settings to config for future use
        self.config.set('wb_sheet', sheet_name)
        self.config.set('date_cell', date_cell)
        self.config.set('start_cell', start_cell)
        self.config.set('end_cell', end_cell)
        self.config.set('duration_cell', duration_cell)
        self.config.set('date_based', date_based)

        # Export to Excel
        export_to_excel(self.excel_file, sheet_name, date_cell, start_cell, end_cell, duration_cell, date_based, self.db)

        self.accept()  # Close the dialog after saving
