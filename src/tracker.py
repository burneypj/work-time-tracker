from PyQt5 import QtWidgets, QtCore
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QSystemTrayIcon, QMenu
import datetime
from db import WorkSessionDB
from exporter import export_to_excel
from config import Config
import threading
import time
import ctypes
import ctypes.wintypes
import win32con
import win32gui
import win32ts
from db import WorkSessionDB
from exporter import handle_excel_export


class TimeTrackerApp(QtWidgets.QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db  # WorkSessionDB instance for database access
        self.start_time = None
        self.end_time = None
        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.tray_icon = None
        self.idle_threshold = 300  # 5 minutes
        self.session_was_stopped_due_to_idle = False
        self.session_was_stopped_due_to_lock = False
        self.start_inactivity_monitor()
        self.register_session_monitor()

        # UI elements
        self.init_ui()
        self.init_tray_icon()

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

    def init_tray_icon(self):
        """Initialize the system tray icon and menu."""
        self.tray_icon = QSystemTrayIcon(QIcon("resources/tray.ico"), self)
        tray_menu = QMenu()

        start_action = tray_menu.addAction("Start Session")
        stop_action = tray_menu.addAction("Stop Session")
        restore_action = tray_menu.addAction("Restore")

        start_action.triggered.connect(self.start_session)
        stop_action.triggered.connect(self.stop_session)
        restore_action.triggered.connect(self.restore_from_tray)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_icon_activated)
        self.tray_icon.show()

    def minimize_to_tray(self):
        """Minimize the application to the system tray."""
        self.hide()

    def restore_from_tray(self):
        """Restore the application from the system tray."""
        self.show()
        self.raise_()

    def on_tray_icon_activated(self, reason):
        """Handle tray icon activation."""
        if reason == QSystemTrayIcon.Trigger:
            self.restore_from_tray()

    def changeEvent(self, event):
        """Override change event to minimize to tray when minimized."""
        if event.type() == QtCore.QEvent.WindowStateChange and self.isMinimized():
            self.minimize_to_tray()
            event.ignore()
        else:
            super().changeEvent(event)

    @QtCore.pyqtSlot()
    def start_session(self):
        """ Start the session """
        self.start_time = datetime.datetime.now()
        self.timer.start(1000)  # Update every second
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)

    @QtCore.pyqtSlot()
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

    def get_idle_duration(self):
        """Get the duration of user inactivity in seconds."""
        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.wintypes.UINT), ("dwTime", ctypes.wintypes.DWORD)]

        lii = LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(lii)
        if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
            millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
            return millis / 1000.0
        return 0

    def start_inactivity_monitor(self):
        """Start a thread to monitor user inactivity."""
        def monitor():
            while True:
                idle_time = self.get_idle_duration()
                if idle_time >= self.idle_threshold:
                    if self.start_time is not None and not self.session_was_stopped_due_to_idle:
                        QtCore.QMetaObject.invokeMethod(self, "stop_session", QtCore.Qt.QueuedConnection)
                        self.session_was_stopped_due_to_idle = True
                elif idle_time < 2:
                    if self.session_was_stopped_due_to_idle:
                        QtCore.QMetaObject.invokeMethod(self, "start_session", QtCore.Qt.QueuedConnection)
                        self.session_was_stopped_due_to_idle = False
                time.sleep(5)

        threading.Thread(target=monitor, daemon=True).start()

    def register_session_monitor(self):
        """Register a hidden window to monitor Windows lock/unlock events."""
        def monitor():
            WNDPROCTYPE = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_int, ctypes.c_uint, ctypes.c_int, ctypes.c_int)

            def wnd_proc(hwnd, msg, wparam, lparam):
                if msg == 0x02B1:  # WM_WTSSESSION_CHANGE
                    if wparam == 0x7:  # WTS_SESSION_LOCK
                        if self.start_time is not None:
                            QtCore.QMetaObject.invokeMethod(self, "stop_session", QtCore.Qt.QueuedConnection)
                            self.session_was_stopped_due_to_lock = True
                    elif wparam == 0x8:  # WTS_SESSION_UNLOCK
                        if self.session_was_stopped_due_to_lock:
                            QtCore.QMetaObject.invokeMethod(self, "start_session", QtCore.Qt.QueuedConnection)
                            self.session_was_stopped_due_to_lock = False

                elif msg == win32con.WM_QUERYENDSESSION:
                    # System is shutting down or user is logging off
                    QtCore.QMetaObject.invokeMethod(self, "exit_app", QtCore.Qt.QueuedConnection)
                    return True  # Allow shutdown to continue

                elif msg == win32con.WM_ENDSESSION:
                    return 0

                elif msg == win32con.WM_DESTROY:
                    ctypes.windll.wtsapi32.WTSUnRegisterSessionNotification(hwnd)
                    ctypes.windll.user32.PostQuitMessage(0)

                return ctypes.windll.user32.DefWindowProcW(hwnd, msg, wparam, lparam)

            hInstance = ctypes.windll.kernel32.GetModuleHandleW(None)
            className = "HiddenWindowClass_TimeTracker"

            wndClass = win32gui.WNDCLASS()
            wndClass.lpfnWndProc = WNDPROCTYPE(wnd_proc)
            wndClass.hInstance = hInstance
            wndClass.lpszClassName = className
            try:
                win32gui.RegisterClass(wndClass)
            except Exception:
                pass

            hwnd = win32gui.CreateWindow(className, className, 0, 0, 0, 0, 0, 0, 0, hInstance, None)
            ctypes.windll.wtsapi32.WTSRegisterSessionNotification(hwnd, win32ts.NOTIFY_FOR_THIS_SESSION)
            win32gui.PumpMessages()

        threading.Thread(target=monitor, daemon=True).start()


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
