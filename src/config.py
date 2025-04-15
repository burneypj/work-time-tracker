import os
import json

class Config:
    #DB_PATH = os.path.join(os.path.expanduser("~"), "time_tracker.db")
    CONFIG_FILE = os.path.join(os.path.expanduser("~"), "time_tracker_config.json")

    def __init__(self):
        # Initialize settings with default values
        self.settings = {
            "wb_sheet": "Sheet1",
            "date_cell": "A1",
            "start_cell": "B1",
            "end_cell": "C1",
            "duration_cell": "D1",
            "date_based_export": True,
            "excel_path": '',
            "db_path": '',
            "minimized": False
        }
        self.load()

    def set(self, key, value):
        """ Set the configuration key to the given value """
        self.settings[key] = value
        self.save()

    def get(self, key, default=None):
        """ Get the configuration value for the key, or return default if key does not exist """
        return self.settings.get(key, default)

    def save(self):
        """ Save the configuration to a JSON file """
        try:
            with open(self.CONFIG_FILE, 'w') as config_file:
                json.dump(self.settings, config_file, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")

    def load(self):
        """ Load configuration from a JSON file """
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, 'r') as config_file:
                    self.settings.update(json.load(config_file))
            except Exception as e:
                print(f"Error loading config: {e}")

    def delete(self):
        """ Delete the configuration file """
        if os.path.exists(self.CONFIG_FILE):
            try:
                os.remove(self.CONFIG_FILE)
            except Exception as e:
                print(f"Error deleting config: {e}")
