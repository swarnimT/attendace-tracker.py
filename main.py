import sqlite3
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

DB_NAME = "attendance.db"
REPORT_FOLDER = "reports"

os.makedirs(REPORT_FOLDER, exist_ok=True)

connection = sqlite3.connect(DB_NAME)
cursor = connection.cursor()

cursor.execute("""
    CREATE TABLE IF NOT EXISTS students (student_id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, creation_date TEXT NOT NULL)""")

cursor.execute("""
    CREATE TABLE IF NOT EXISTS attendance (attendance_id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER NOT NULL,student_id INTEGER NOT NULL, date TEXT NOT NULL, status TEXT NOT NULL, FOREIGN KEY (student_id) REFERENCES students (student_id),FOREIGN KEY (student_id) REFERENCES students (student_id))""")

sqlite3.Cursor.execute("""
    CREATE TABLE IF NOT EXISTS lecture_summary (summary_id INTEGER PRIMARY KEY AUTOINCREMENT,student_id INTEGER NOT NULL, student_id INTEGER NOT NULL,total_lectures INTEGER NOT NULL,attended_lectures INTEGER NOT NULL, FOREIGN KEY (student_id) REFERENCES students (student_id),FOREIGN KEY (subject_id) REFERENCES subjects (subject_id))""")