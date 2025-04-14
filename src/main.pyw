import sys
import os
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from tracker import TimeTrackerApp
from db import WorkSessionDB
from config import Config


def select_database_file(config):
    """Prompt the user to select a database file or use the existing one from config."""
    db_path = config.get("db_path")

    # If no database path is configured, prompt the user to select one
    if not db_path or not os.path.exists(db_path):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        db_path, _ = QFileDialog.getSaveFileName(
            None,
            "Select Database File",
            "sessions.db",  # Default file name
            "SQLite Database Files (*.db);;All Files (*)",
            options=options
        )

        if not db_path:
            QMessageBox.critical(None, "Error", "No database file selected. Exiting application.")
            sys.exit(1)  # Exit if no file is selected

        # Save the selected database path to the config
        config.set("db_path", db_path)

    return db_path


def main():
    # Initialize the QApplication instance
    app = QtWidgets.QApplication(sys.argv)

    # Load the configuration
    config = Config()

    # Prompt the user to select a database file
    db_path = select_database_file(config)

    # Set up the database
    db = WorkSessionDB(db_path)  # Create a WorkSessionDB instance with the provided path

    # Create the main application window, passing the db object
    window = TimeTrackerApp(db, config)

    # Show the window
    window.show()

    # Start the application event loop
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
