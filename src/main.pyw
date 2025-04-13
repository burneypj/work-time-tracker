import sys
from PyQt5 import QtWidgets
from tracker import TimeTrackerApp
from db import WorkSessionDB

def main():
    # Set up the database
    db_path = "sessions.db"  # Path where your database will be saved
    db = WorkSessionDB(db_path)  # Create a WorkSessionDB instance with the provided path

    # Initialize the QApplication instance
    app = QtWidgets.QApplication(sys.argv)

    # Create the main application window, passing the db object
    window = TimeTrackerApp(db)

    # Show the window
    window.show()

    # Start the application event loop
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
