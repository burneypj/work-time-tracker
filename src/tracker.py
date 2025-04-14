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
        """ Start the session """
        self.start_time = datetime.datetime.now()
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
        """ Update the UI with the current time """
        if self.start_time:
            elapsed_time = datetime.datetime.now() - self.start_time
            hours, remainder = divmod(elapsed_time.total_seconds(), 3600)
            minutes, seconds = divmod(remainder, 60)
            self.duration_label.setText(f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}")

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
                    handle_excel_export(self.db)

    def closeEvent(self, event):
        """Handle the close button (X) to call exit_app."""
        self.exit_app()
        event.accept()  # Accept the event to close the application
