### README.md

# work-time-tracker

The **work-time-tracker (wtt)** is a desktop application built using Python and PyQt5 to help users track their work sessions, export session data to Excel, and manage session data efficiently. It includes features like session monitoring, inactivity detection, and exporting data with customizable configurations.

---

## Features

1. **Session Tracking**:
   - Start and stop work sessions.
   - Automatically log session start and end times, along with the duration.

2. **Inactivity Monitoring**:
   - Automatically stop sessions when the user is idle for a configurable threshold.
   - Resume sessions when activity is detected.

3. **Windows Lock/Unlock Detection**:
   - Automatically stop sessions when the system is locked.
   - Resume sessions when the system is unlocked.

4. **Export to Excel**:
   - Export session data to an Excel file.
   - Configure export settings, including sheet name, starting date, and cell mappings.
   - Supports date-based and flat data exports.

5. **System Tray Integration**:
   - Minimize the application to the system tray.
   - Start/stop sessions and restore the application from the tray.

6. **Configuration Management**:
   - Save and load user preferences, such as export settings and database paths, using a JSON configuration file.

7. **Reset Functionality**:
   - A **Reset** button allows users to delete all session data and configuration files.
   - Displays a confirmation dialog before proceeding.
   - Closes the application after resetting.

---

## File Structure

### 1. **`main.pyw`**
   - Entry point of the application.
   - Initializes the application, loads the configuration, and sets up the database.

### 2. **`tracker.py`**
   - Contains the `TimeTrackerApp` class, which manages the main application logic.
   - Handles session tracking, inactivity monitoring, system tray integration, and reset functionality.

### 3. **`db.py`**
   - Contains the `WorkSessionDB` class for managing the SQLite database.
   - Provides methods to add, retrieve, and filter session data.

### 4. **`exporter.py`**
   - Contains the `ExportConfigDialog` class for configuring and exporting session data to Excel.
   - Supports customizable export settings and date-based formatting.

### 5. **`config.py`**
   - Contains the `Config` class for managing user preferences.
   - Saves and loads settings from a JSON file.

### 6. **`utils.py`**
   - Utility functions for formatting durations and incrementing Excel cell references.

---

## Installation

### Prerequisites
- Python 3.8 or higher
- Required Python packages:
  - `PyQt5`
  - `xlwings`
  - `pywin32`

### Steps
1. Clone the repository:
   ```bash
   git clone https:https://github.com/burneypj/work-time-tracker.git
   cd work-time-tracker
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python src/main.pyw
   ```

---

## Usage

### Starting the Application
- On the first run, you will be prompted to select or create a database file for storing session data.

### Tracking Sessions
1. Click **Start** to begin a session.
2. Click **Stop** to end the session and save it to the database.

### Exporting Data
1. Click **Export to Excel**.
2. Configure the export settings, including the starting date and cell mappings.
3. Click **Save** to export the data to the selected Excel file.

### System Tray
- Minimize the application to the tray for background operation.
- Use the tray menu to start/stop sessions or restore the application.

---

## Configuration

The application saves user preferences in a JSON file located at:
```
~\time_tracker_config.json
```

### Configurable Settings
- `wb_sheet`: Default Excel sheet name.
- `date_cell`, `start_cell`, `end_cell`, `duration_cell`: Default cell mappings for export.
- `date_based_export`: Whether to use date-based export formatting.
- `excel_path`: Path to the last used Excel file.
- `db_path`: Path to the database file.
- `minimized`: Whether the application starts minimized.

---

## Database Schema

The SQLite database contains a single table:

### `work_sessions`
| Column      | Type    | Description                     |
|-------------|---------|---------------------------------|
| `id`        | INTEGER | Primary key (auto-increment).   |
| `start_time`| TEXT    | Session start time (ISO format).|
| `end_time`  | TEXT    | Session end time (ISO format).  |
| `duration`  | TEXT    | Session duration (in seconds).  |

---

## Utility Functions

### utils.py
- **`increment_cell_row(cell_ref)`**:
  - Increments the row number in an Excel cell reference (e.g., `A1` → `A2`).
- **`format_duration(seconds)`**:
  - Converts a duration in seconds to `HH:MM:SS` format.

---

## Future Enhancements
- Add support for weekly and monthly summary reports.
- Enable cloud synchronization for session data.
- Add more export formats (e.g., CSV, PDF).

---

## Contributing
Contributions are welcome! Please open an issue or submit a pull request for any enhancements or bug fixes.

---

## License
This project is licensed under the MIT License. See the LICENSE file for more details.

---

## Acknowledgments
- **PyQt5** for the GUI framework.
- **xlwings** for Excel integration.
- **pywin32** for Windows-specific functionalities.
