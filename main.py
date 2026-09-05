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
    CREATE TABLE IF NOT EXISTS attendance (attendance_id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER NOT NULL,student_id INTEGER NOT NULL, date TEXT NOT NULL, status TEXT NOT NULL, FOREIGN KEY (student_id) REFERENCES students (student_id),FOREIGN KEY (subject_id) REFERENCES subjects (subject_id))""")

cursor.execute("""
    CREATE TABLE IF NOT EXISTS lecture_summary (summary_id INTEGER PRIMARY KEY AUTOINCREMENT,student_id INTEGER NOT NULL, student_id INTEGER NOT NULL,total_lectures INTEGER NOT NULL,attended_lectures INTEGER NOT NULL, FOREIGN KEY (student_id) REFERENCES students (student_id),FOREIGN KEY (subject_id) REFERENCES subjects (subject_id))""")

connection.commit()

def get_students():
    cursor.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
    return cursor.fetchall()

def get_subjects():
    cursor.execute("SELECT * FROM subjects WHERE student_id=?", (student_id,))
    return cursor.fetchall()


def clean_filename(name):
    return "".join(c for c in name if c.isalnum() or c in (" ", "_", "-")).strip().replace(" ", "_")

def create_student():
    print("\n Create Student")
    name=input("Enter student name: ").strip()
    if not name:
        print("Student name cannot be empty.")
        return
    creation_date = input("Enter creation date (DD-MM-YYYY): ").strip()
    try:
        datetime.strptime(creation_date, "%d-%m-%Y")
    except ValueError:
        print("Invalid date format. Please use DD-MM-YYYY.")
        return
    cursor.execute("""INSERT INTO students (name, creation_date) VALUES (?, ?)""", (name, creation_date))
    student_id = cursor.lastrowid
    connection.commit()
    print("/nHow many subjects does the you have?")

    try:
        number_of_subjects = int(input("Enter number of subjects: "))
    except ValueError:
        print("Invalid input. Please enter a valid number.")
        return
    if number_of_subjects <= 0:
        print("you need atleast one subject.")
        return
    print("\nAdd subjects")
    for i in range(number_of_subjects):
        subject_name = input(f"Enter subject {i + 1} name: ").strip()
        if not subject_name:
            print("Subject name cannot be empty.")
            return
        cursor.execute("""INSERT INTO subjects (student_id, name) VALUES (?, ?)""", (student_id, subject_name))