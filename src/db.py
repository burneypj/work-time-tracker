import sqlite3
import os

class WorkSessionDB:
    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()
        self.create_table()

    def create_table(self):
        """Create a table for storing work sessions if it doesn't already exist."""
        query = '''CREATE TABLE IF NOT EXISTS work_sessions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        start_time TEXT NOT NULL,
                        end_time TEXT NOT NULL,
                        duration TEXT NOT NULL
                    )'''
        self.cursor.execute(query)
        self.conn.commit()

    def add_session(self, start_time, end_time, duration):
        """Add a new session to the database."""
        query = '''INSERT INTO work_sessions (start_time, end_time, duration)
                   VALUES (?, ?, ?)'''
        self.cursor.execute(query, (start_time, end_time, duration))
        self.conn.commit()

    def get_sessions(self, start_date=None):
        """Retrieve work sessions from the database, optionally starting from a specific date."""
        if start_date:
            query = '''SELECT start_time, end_time, duration
                       FROM work_sessions
                       WHERE DATE(start_time) >= ?'''
            self.cursor.execute(query, (start_date.isoformat(),))
        else:
            query = '''SELECT start_time, end_time, duration FROM work_sessions'''
            self.cursor.execute(query)
        return self.cursor.fetchall()

    def get_last_session(self):
        """Retrieve the last saved work session from the database."""
        query = '''SELECT start_time, end_time, duration
                   FROM work_sessions
                   ORDER BY id DESC
                   LIMIT 1'''
        self.cursor.execute(query)
        return self.cursor.fetchone()

    def delete(self):
        """Delete the database file."""
        self.close()
        if os.path.exists(self.db_path):
            try:
                os.remove(self.db_path)
            except Exception as e:
                print(f"Error deleting database: {e}")

    def close(self):
        """Close the database connection."""
        self.conn.close()
