from PyQt5 import QtWidgets, QtCore
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QSystemTrayIcon, QMenu
import datetime
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
    def __init__(self, db, cfg):
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
        # Check if the app shall start minimized
        self.cfg = cfg
        # Check for daily limit in config
        self.daily_limit = self.cfg.get("daily_limit")
        if self.daily_limit is None:
            self.daily_limit = self.prompt_for_daily_limit()  # Ask the user for the daily limit
            self.cfg.set("daily_limit", self.daily_limit)  # Save it to the config
        if self.cfg.get("minimized", True):
            QtCore.QTimer.singleShot(0, self.minimize_to_tray)
            # start session automatically on minimized startup
            self.start_session()

    def init_ui(self):
        # Set up the UI for session tracking
        self.setWindowTitle('Time Tracker')
        self.setGeometry(100, 100, 300, 200)

        self.month_change_checked = False

        self.start_button = QtWidgets.QPushButton('Start', self)
        self.stop_button = QtWidgets.QPushButton('Stop', self)
        self.export_button = QtWidgets.QPushButton('Export to Excel', self)
        self.reset_button = QtWidgets.QPushButton('Reset', self)

        # Add a label to display the running duration
        self.duration_label = QtWidgets.QLabel("Duration: 00:00:00", self)
        self.duration_label.setAlignment(QtCore.Qt.AlignCenter)

        self.start_button.clicked.connect(self.start_session)
        self.stop_button.clicked.connect(self.stop_session)
        self.export_button.clicked.connect(self.export_to_excel)
        self.reset_button.clicked.connect(self.reset_app)
        self.reset_button.setStyleSheet("font-size: 12px; color: red;")
        self.reset_button.setFixedSize(80, 30)

        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)

        # Layout setup
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self.duration_label)  # Add the duration label to the layout
        layout.addWidget(self.start_button)
        layout.addWidget(self.stop_button)
        layout.addWidget(self.export_button)
        layout.addWidget(self.reset_button)  # Add the Reset button to the layout
        self.setLayout(layout)

    def export_to_excel(self):
        """Handle exporting session data to Excel."""
        handle_excel_export(self.db, self.cfg)

    def update_tray_menu(self):
        if self.tray_icon:
            tray_menu = QMenu()
            start_action = tray_menu.addAction("Start Session")
            start_action.setEnabled(self.start_button.isEnabled())
            stop_action = tray_menu.addAction("Stop Session")
            stop_action.setEnabled(self.stop_button.isEnabled())
            restore_action = tray_menu.addAction("Restore")
            Close_action = tray_menu.addAction("Close")

            start_action.triggered.connect(self.start_session)
            stop_action.triggered.connect(self.stop_session)
            restore_action.triggered.connect(self.restore_from_tray)
            Close_action.triggered.connect(self.exit_app)

            self.tray_icon.setContextMenu(tray_menu)

    def minimize_to_tray(self):
        """Minimize the application to the system tray."""
        self.hide()
        self.cfg.set("minimized", True)
        if not self.tray_icon:
            self.tray_icon = QSystemTrayIcon(QIcon("resources/tray.ico"), self)
            self.tray_icon.setToolTip("Time Tracker - Minimized to Tray")
            self.update_tray_menu()
            self.tray_icon.activated.connect(self.on_tray_icon_activated)
        self.tray_icon.show()

    def restore_from_tray(self):
        """Restore the application from the system tray."""
        self.show()
        self.raise_()
        self.activateWindow()
        self.showNormal()
        if self.tray_icon:
            self.tray_icon.hide()
        self.cfg.set("minimized", False)

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
        """Start the session."""
        self.start_time = datetime.datetime.now()
        self.total_time_today = self.get_total_time_today()  # Calculate total time for the day
        self.daily_limit_exceeded = False  # Reset the daily limit flag
        self.timer.start(1000)
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.update_tray_menu()
        if not self.month_change_checked:
            self.check_Month_Change()
            self.month_change_checked = True

    @QtCore.pyqtSlot()
    def stop_session(self):
        """ Stop the session and save data """
        if self.start_time is not None:
            self.end_time = datetime.datetime.now()
            total_seconds = int((self.end_time - self.start_time).total_seconds())
            # Log the session to the database (store duration in seconds)
            self.db.add_session(self.start_time, self.end_time, total_seconds)
            # Stop the timer and reset buttons
            self.timer.stop()
            self.start_time = None
            self.start_button.setEnabled(True)
            self.stop_button.setEnabled(False)
            self.update_tray_menu()

    @QtCore.pyqtSlot()
    def exit_app(self):
        """Exit the application gracefully."""
        # Stop the session
        self.stop_session()
        # Close the tray icon if it exists
        if self.tray_icon:
            self.tray_icon.hide()
        # Close the application
        QtWidgets.QApplication.quit()

    def update_time(self):
        """Update the UI with the current time and check daily limit."""
        if self.start_time:
            elapsed_time = datetime.datetime.now() - self.start_time
            hours, remainder = divmod(elapsed_time.total_seconds(), 3600)
            minutes, seconds = divmod(remainder, 60)
            self.duration_label.setText(f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}")
            self.check_daily_limit()  # Check if the daily limit is exceeded

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

    def check_Month_Change(self):
        """Check if the month has changed since the last session."""
        sessions = self.db.get_last_session()
        if sessions:
            # Extract last session month
            last_session_month = datetime.datetime.fromisoformat(sessions[0]).month
            # Check if the month has changed
            if last_session_month != datetime.datetime.now().month:
                # Open a popup with an Export to Excel button
                msg_box = QtWidgets.QMessageBox(self)
                msg_box.setIcon(QtWidgets.QMessageBox.Information)
                msg_box.setWindowTitle("Month Change Detected")
                msg_box.setText("The month has changed. Would you like to export the time data?")
                export_button = msg_box.addButton("Export to Excel", QtWidgets.QMessageBox.AcceptRole)
                msg_box.addButton("Cancel", QtWidgets.QMessageBox.RejectRole)
                msg_box.exec_()
                if msg_box.clickedButton() == export_button:
                    handle_excel_export(self.db, self.cfg)

    def closeEvent(self, event):
        """Handle the close button (X) to call exit_app."""
        self.exit_app()
        event.accept()  # Accept the event to close the application

    def reset_app(self):
        """Reset the application by deleting the session database and config file."""
        reply = QtWidgets.QMessageBox.question(
            self,
            "Confirm Reset",
            "This will delete all session data and configurations. Are you sure?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No
        )
        if reply == QtWidgets.QMessageBox.Yes:
            # stop running timers
            self.timer.stop()
            # delete the database connection
            self.db.delete()  # Ensure the database is closed before deletion
            # Delete the configuration file
            self.cfg.delete()
            # Exit the application
            QtWidgets.QMessageBox.information(self, "Reset Complete", "The application will now exit.")
            QtWidgets.QApplication.quit()

    def get_total_time_today(self):
        """Calculate the total session time for the current day."""
        today = datetime.datetime.now().date()
        sessions = self.db.get_sessions(today)
        total_seconds = sum(int(session[2]) for session in sessions)
        return total_seconds

    def check_daily_limit(self):
        """Check if the total duration for the day exceeds 8 hours."""
        if self.start_time:
            elapsed_time = datetime.datetime.now() - self.start_time
            elapsed_seconds = elapsed_time.total_seconds()
        else:
            elapsed_seconds = 0

        total_time_today = self.total_time_today + elapsed_seconds
        if total_time_today >= self.daily_limit and not self.daily_limit_exceeded:
            self.daily_limit_exceeded = True
            hours, remainder = divmod(total_time_today, 3600)
            minutes, _ = divmod(remainder, 60)
        # Show a system tray notification instead of a message box
            if self.tray_icon:
                self.tray_icon.showMessage(
                    "Daily Limit Exceeded",
                    f"Warning: You have worked for {int(hours)} hours and {int(minutes)} minutes today. "
                    "Consider taking a break!",
                    QSystemTrayIcon.Warning,
                    5000  # Duration in milliseconds
                )
            else:
                msg_box = QtWidgets.QMessageBox()
                msg_box.setIcon(QtWidgets.QMessageBox.Warning)
                msg_box.setWindowTitle("Daily Limit Exceeded")
                msg_box.setText(
                    f"Warning: You have worked for {int(hours)} hours and {int(minutes)} minutes today. "
                    "Consider taking a break!"
                )
                msg_box.addButton("OK", QtWidgets.QMessageBox.AcceptRole)
                msg_box.exec_()
            # Add 30 minutes to the daily limit and reset the flag
            self.daily_limit += 1800  # Increase limit by 30 minutes
            self.daily_limit_exceeded = False  # Reset the flag for the next check

    def prompt_for_daily_limit(self):
        """Prompt the user to input the daily limit in hours and minutes."""
        while True:
            # Create a custom dialog for entering hours and minutes
            dialog = QtWidgets.QDialog(self)
            dialog.setWindowTitle("Set Daily Limit")
            dialog.setModal(True)

            layout = QtWidgets.QVBoxLayout(dialog)

            # Input fields for hours and minutes
            hours_label = QtWidgets.QLabel("Enter hours:")
            hours_input = QtWidgets.QSpinBox()
            hours_input.setRange(0, 24)
            hours_input.setValue(8)  # Default value

            minutes_label = QtWidgets.QLabel("Enter minutes:")
            minutes_input = QtWidgets.QSpinBox()
            minutes_input.setRange(0, 59)
            minutes_input.setValue(0)  # Default value

            layout.addWidget(hours_label)
            layout.addWidget(hours_input)
            layout.addWidget(minutes_label)
            layout.addWidget(minutes_input)

            # OK and Cancel buttons
            button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
            layout.addWidget(button_box)

            button_box.accepted.connect(dialog.accept)
            button_box.rejected.connect(dialog.reject)

            # Show the dialog
            if dialog.exec_() == QtWidgets.QDialog.Accepted:
                # Get the values from the input fields
                hours = hours_input.value()
                minutes = minutes_input.value()
                if hours == 0 and minutes == 0:
                    QtWidgets.QMessageBox.warning(
                        self,
                        "Invalid Input",
                        "You must set a daily limit greater than 0."
                    )
                    continue
                # Convert hours and minutes to seconds and return
                return (hours * 3600) + (minutes * 60)
            else:
                # If the user cancels, show a warning and retry
                QtWidgets.QMessageBox.warning(
                    self,
                    "Daily Limit Required",
                    "You must set a daily limit to proceed."
                )
